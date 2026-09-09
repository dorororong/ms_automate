"""Local SQLite mirror of the current Outlook profile.

Duplicate checking used to read every Calendar and Task item over COM right
before the cards were shown and again right before saving.  That is fine for
one batch per message, but the review window now registers one card at a
time, and a full COM scan on every click is too slow.

So the profile is mirrored into SQLite instead:

    앱 시작        → 백그라운드 전체 동기화 (전용 COM 스레드)
    이후           → 주기적 재동기화 + 검토창을 열 때 갱신 요청
    이 앱의 등록   → Outlook과 미러에 함께 기록 (write-through)

The one thing a mirror cannot do is see an item another program created
since the last sync.  `is_stale()` reports that window so the UI can say so.
"""

from __future__ import annotations

import sqlite3
import threading
from datetime import datetime, timedelta, timezone
from typing import Any, Callable

from models import CALENDAR, DUE_MARKER_CATEGORY, TODO, WorkItem, due_marker_subject
from outlook_adapter import (
    OutlookAdapter,
    OutlookEntry,
    OutlookOperationError,
    OutlookUnavailableError,
    parse_outlook_datetime,
    parse_outlook_due,
)
from storage import Database, utc_now_precise_iso

try:  # The sync thread needs its own COM apartment.
    import pythoncom
except ImportError:  # pragma: no cover - non-Windows or missing pywin32
    pythoncom = None


SYNC_INTERVAL_SECONDS = 300.0
STALE_AFTER_SECONDS = 300.0
MIRROR_READ_LIMIT = 20000


def _iso(value: datetime | None) -> str | None:
    return None if value is None else value.isoformat(timespec="seconds")


def _parse_iso(value: Any) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.fromisoformat(str(value))
    except ValueError:
        return None


def entry_to_row(entry: OutlookEntry) -> dict[str, Any]:
    """Flatten one Outlook item into a mirror row."""

    return {
        "entry_id": entry.entry_id,
        "item_type": entry.item_type,
        "subject": entry.subject,
        "start": _iso(entry.start),
        "end": _iso(entry.end),
        "due": _iso(entry.due),
        "all_day": 1 if entry.all_day else 0,
        "location": entry.location or "",
        "category": entry.category or "",
        "reminder": _iso(entry.reminder),
        "reminder_minutes": entry.reminder_minutes,
        "complete": None if entry.complete is None else int(entry.complete),
        "body": entry.body or "",
    }


def row_to_entry(row: sqlite3.Row | dict[str, Any]) -> OutlookEntry:
    data = dict(row)
    complete = data.get("complete")
    return OutlookEntry(
        item_type=str(data["item_type"]),
        entry_id=str(data["entry_id"]),
        subject=str(data.get("subject") or ""),
        start=_parse_iso(data.get("start")),
        end=_parse_iso(data.get("end")),
        due=_parse_iso(data.get("due")),
        all_day=bool(data.get("all_day")),
        location=str(data.get("location") or ""),
        category=str(data.get("category") or ""),
        reminder=_parse_iso(data.get("reminder")),
        reminder_minutes=(
            None
            if data.get("reminder_minutes") is None
            else int(data["reminder_minutes"])
        ),
        complete=None if complete is None else bool(complete),
        body=str(data.get("body") or ""),
    )


def work_item_to_row(item: WorkItem) -> dict[str, Any] | None:
    """Build the mirror row for an item this app just wrote to Outlook.

    Returns None when the item never reached Outlook, so the caller does not
    have to special-case reference-only rows.
    """

    entry_id = (item.outlook_entry_id or "").strip()
    target = item.outlook_target
    if not entry_id or target not in (CALENDAR, TODO):
        return None

    start = end = due = None
    if target == CALENDAR:
        try:
            start = parse_outlook_datetime(item.start) if item.start else None
        except ValueError:
            start = None
        try:
            end = parse_outlook_datetime(item.end) if item.end else None
        except ValueError:
            end = None
    else:
        try:
            due = parse_outlook_due(item.due) if item.due else None
        except ValueError:
            due = None

    return {
        "entry_id": entry_id,
        "item_type": target,
        "subject": item.title,
        "start": _iso(start),
        "end": _iso(end),
        "due": _iso(due),
        "all_day": 1 if item.all_day else 0,
        "location": item.location or "",
        "category": item.category or "",
        "reminder": None,
        "reminder_minutes": item.reminder_minutes,
        "complete": 0 if target == TODO else None,
        "body": item.description or "",
    }


def due_marker_to_row(item: WorkItem) -> dict[str, Any] | None:
    """마감 표시(종일 일정)의 미러 행. 작업 행과는 별개의 Outlook 항목이다."""

    entry_id = (item.calendar_entry_id or "").strip()
    if not entry_id or not item.due:
        return None
    try:
        due = parse_outlook_due(item.due)
    except ValueError:
        return None
    start = datetime.combine(due.date(), datetime.min.time())
    categories = [DUE_MARKER_CATEGORY]
    if item.category:
        categories.append(item.category)
    return {
        "entry_id": entry_id,
        "item_type": CALENDAR,
        "subject": due_marker_subject(item.title),
        "start": _iso(start),
        "end": _iso(start + timedelta(days=1)),
        "due": None,
        "all_day": 1,
        "location": item.location or "",
        "category": ", ".join(categories),
        "reminder": None,
        "reminder_minutes": None,
        "complete": None,
        "body": item.description or "",
    }


class OutlookMirror:
    """Reads the mirror for callers; keeps it fresh in a background thread."""

    def __init__(
        self,
        database: Database,
        *,
        adapter_factory: Callable[[], OutlookAdapter] = OutlookAdapter,
        interval_seconds: float = SYNC_INTERVAL_SECONDS,
    ) -> None:
        self.database = database
        self.adapter_factory = adapter_factory
        self.interval_seconds = interval_seconds
        self._thread: threading.Thread | None = None
        self._wake = threading.Event()
        self._stop = threading.Event()
        self._on_status: Callable[[str, object], None] | None = None
        self._last_error: str | None = None

    # --- 읽기 ----------------------------------------------------------

    def entries(self) -> list[OutlookEntry]:
        rows = self.database.read_outlook_mirror()
        return [row_to_entry(row) for row in rows[:MIRROR_READ_LIMIT]]

    def last_synced_at(self) -> datetime | None:
        raw = self.database.get_sync_state("outlook_mirror_synced_at")
        if not raw:
            return None
        try:
            return datetime.fromisoformat(raw.replace("Z", "+00:00"))
        except ValueError:
            return None

    def has_snapshot(self) -> bool:
        return self.last_synced_at() is not None

    def age_seconds(self) -> float | None:
        synced = self.last_synced_at()
        if synced is None:
            return None
        return (datetime.now(timezone.utc) - synced).total_seconds()

    def is_stale(self, limit_seconds: float = STALE_AFTER_SECONDS) -> bool:
        age = self.age_seconds()
        return age is None or age > limit_seconds

    @property
    def last_error(self) -> str | None:
        return self._last_error

    # --- 쓰기 (write-through) -------------------------------------------

    def write_through(self, item: WorkItem) -> None:
        """Record an item this app just created so the next check sees it."""

        row = work_item_to_row(item)
        if row is None:
            return
        try:
            self.database.upsert_outlook_entry(row)
        except sqlite3.Error:
            # 미러는 캐시다. 기록에 실패해도 등록 자체는 이미 끝났다.
            self.request_refresh()

    def write_through_due_marker(self, item: WorkItem) -> None:
        """마감 표시도 미러에 넣어 같은 표시를 두 번 만들지 않게 한다."""

        row = due_marker_to_row(item)
        if row is None:
            return
        try:
            self.database.upsert_outlook_entry(row)
        except sqlite3.Error:
            self.request_refresh()

    def forget(self, entry_id: str) -> None:
        if not entry_id:
            return
        try:
            self.database.delete_outlook_entry(entry_id)
        except sqlite3.Error:
            self.request_refresh()

    # --- 동기화 ---------------------------------------------------------

    def sync_once(self, adapter: OutlookAdapter | None = None) -> int:
        """Read the whole profile and replace the mirror. Raises on failure."""

        started_at = utc_now_precise_iso()
        source = adapter or self.adapter_factory()
        entries = source.read_entries(limit=MIRROR_READ_LIMIT)
        rows = [entry_to_row(entry) for entry in entries if entry.entry_id]
        return self.database.sync_outlook_mirror(rows, started_at=started_at)

    def start(self, on_status: Callable[[str, object], None] | None = None) -> None:
        if self._thread is not None and self._thread.is_alive():
            return
        self._on_status = on_status
        self._stop.clear()
        self._wake.set()  # 첫 동기화는 기다리지 않고 바로 시작한다.
        self._thread = threading.Thread(
            target=self._worker, name="outlook-mirror-sync", daemon=True
        )
        self._thread.start()

    def request_refresh(self) -> None:
        self._wake.set()

    def refresh_if_stale(self, limit_seconds: float = STALE_AFTER_SECONDS) -> None:
        if self.is_stale(limit_seconds):
            self.request_refresh()

    def stop(self) -> None:
        self._stop.set()
        self._wake.set()

    def _worker(self) -> None:
        """Own COM apartment; the Tk thread keeps its own adapter."""

        if pythoncom is not None:
            pythoncom.CoInitialize()
        adapter: OutlookAdapter | None = None
        try:
            while not self._stop.is_set():
                self._wake.wait(self.interval_seconds)
                if self._stop.is_set():
                    return
                self._wake.clear()
                try:
                    if adapter is None:
                        adapter = self.adapter_factory()
                    count = self.sync_once(adapter)
                except (
                    OutlookUnavailableError,
                    OutlookOperationError,
                    AttributeError,
                    OSError,
                    ValueError,
                    sqlite3.Error,
                ) as exc:
                    adapter = None  # 연결이 끊겼을 수 있으므로 다음 회차에 다시 만든다.
                    self._last_error = str(exc)
                    self._notify("mirror_error", str(exc))
                else:
                    self._last_error = None
                    self._notify("mirror_synced", count)
        finally:
            if pythoncom is not None:
                pythoncom.CoUninitialize()

    def _notify(self, kind: str, payload: object) -> None:
        callback = self._on_status
        if callback is None:
            return
        try:
            callback(kind, payload)
        except Exception:
            # 상태 표시 실패가 동기화 루프를 멈추게 두지 않는다.
            pass
