"""Application service: capture text, classify it, record it in Outlook.

This is the CATMOA pipeline reduced to what this app actually does:

    prepare(텍스트)  → 분류 + SQLite pending 기록
    record(항목들)   → 클래식 Outlook COM 기록 + saved/failed 표시

There is no mail sending and no Microsoft Graph path; Outlook is driven
directly through COM.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
import re

from classifier import ClassificationResult, classify_text
from models import (
    CALENDAR,
    DUE_MARKER_CATEGORY,
    ReviewIssue,
    SCOPE_REFERENCE,
    TODO,
    WorkItem,
    due_marker_subject,
)
from outlook_adapter import (
    OutlookEntry,
    OutlookAdapter,
    OutlookOperationError,
    OutlookUnavailableError,
    parse_outlook_datetime,
)
from outlook_mirror import OutlookMirror
from storage import Database
from upstage_classifier import classify_with_upstage, has_api_key

APP_ROOT = Path(__file__).resolve().parent
DATABASE_PATH = APP_ROOT / "data" / "work_items.sqlite3"


@dataclass
class Analysis:
    """One analysed input, already persisted as pending work items."""

    input_id: int
    text: str
    work_items: list[WorkItem]
    ignored: list[str] = field(default_factory=list)
    source: str = "offline"
    model: str = "rule-based"
    warning: str | None = None
    context: dict[str, object] = field(default_factory=dict)

    @property
    def registerable(self) -> list[WorkItem]:
        """Outlook 대상이 될 수 있는 항목 (참조 제외)."""

        return [
            i
            for i in self.work_items
            if i.outlook_target is not None
            and i.scope != SCOPE_REFERENCE
            and i.source_state not in {"completed", "informational"}
        ]

    @property
    def ready_to_register(self) -> list[WorkItem]:
        """중복 제외까지 통과해 현재 등록 가능한 항목."""

        return [
            i for i in self.registerable
            if i.can_register and not i.excluded_reason
        ]

    def count(self, action: str) -> int:
        return sum(i.action == action for i in self.work_items)

    @property
    def excluded_count(self) -> int:
        return sum(bool(i.excluded_reason) for i in self.work_items)

    @property
    def summary(self) -> str:
        from models import ACTION_TODO, ACTION_WITH_DATE, ACTION_WITH_DUE

        summary = (
            f"바로 할 일 {self.count(ACTION_TODO)} · "
            f"일정 {self.count(ACTION_WITH_DATE)} · "
            f"마감 {self.count(ACTION_WITH_DUE)} · "
            f"참조 {sum(i.scope == SCOPE_REFERENCE for i in self.work_items)}"
        )
        if self.excluded_count:
            summary += f" · 중복 제외 {self.excluded_count}"
        return summary


@dataclass
class RecordOutcome:
    """Result of writing selected work items to Outlook."""

    saved: int = 0
    failed: int = 0
    skipped: int = 0  # 참조 — 로컬 기록만
    errors: dict[int, str] = field(default_factory=dict)  # index → message
    duplicates: dict[int, str] = field(default_factory=dict)  # index → reason
    warning: str | None = None
    blocked: bool = False  # 등록 전 필수 검사 실패


def _normalise_title(value: str) -> str:
    """Compare titles without whitespace or punctuation noise."""

    return re.sub(r"[\W_]+", "", value.casefold(), flags=re.UNICODE)


def _calendar_interval(
    start: datetime | None,
    end: datetime | None,
    *,
    all_day: bool = False,
) -> tuple[datetime, datetime] | None:
    if start is None:
        return None
    if end is None or end <= start:
        end = start + (timedelta(days=1) if all_day else timedelta(minutes=60))
    return start, end


def _draft_interval(item: WorkItem) -> tuple[datetime, datetime] | None:
    if item.outlook_target != CALENDAR or not item.start:
        return None
    start = parse_outlook_datetime(item.start)
    end = parse_outlook_datetime(item.end) if item.end else None
    return _calendar_interval(start, end, all_day=item.all_day)


def _is_due_marker(entry: OutlookEntry) -> bool:
    """이 앱이 만든 마감 표시인가."""

    categories = [part.strip() for part in (entry.category or "").split(",")]
    return DUE_MARKER_CATEGORY in categories


def _entry_interval(entry: OutlookEntry) -> tuple[datetime, datetime] | None:
    if entry.item_type != CALENDAR or _is_due_marker(entry):
        # 마감 표시는 하루를 차지하지만 일정이 아니다. 겹침으로 세면 그날의
        # 실제 일정 등록이 모두 막힌다.
        return None
    return _calendar_interval(entry.start, entry.end, all_day=entry.all_day)


def _overlaps(
    left: tuple[datetime, datetime], right: tuple[datetime, datetime]
) -> bool:
    return left[0] < right[1] and right[0] < left[1]


class WorkflowService:
    def __init__(
        self,
        database: Database | None = None,
        outlook: OutlookAdapter | None = None,
        mirror: OutlookMirror | None = None,
    ) -> None:
        self.database = database or Database(DATABASE_PATH)
        self.outlook = outlook or OutlookAdapter()
        self.mirror = mirror or OutlookMirror(self.database)

    # --- analysis -----------------------------------------------------

    def has_api_key(self) -> bool:
        return has_api_key()

    def classify(
        self,
        text: str,
        *,
        use_ai: bool = True,
        reference_date: str | None = None,
        sent_at: str | None = None,
    ) -> tuple[ClassificationResult, str, str, str | None]:
        """Classify without touching the database.

        Returns (result, source, model, warning). Falls back to the offline
        rules when the AI call fails so the hotkey flow never dead-ends.
        """

        if not text.strip():
            raise ValueError("분석할 텍스트가 비어 있습니다.")
        if use_ai and self.has_api_key():
            try:
                return (
                    classify_with_upstage(
                        text,
                        reference_date=reference_date,
                        sent_at=sent_at,
                    ),
                    "upstage",
                    "solar-pro4",
                    None,
                )
            except Exception as exc:
                warning = f"Solar Pro 4 분석 실패로 오프라인 규칙을 사용했습니다: {exc}"
                return classify_text(text), "offline", "rule-based", warning
        return classify_text(text), "offline", "rule-based", None

    def prepare(
        self,
        text: str,
        *,
        use_ai: bool = True,
        reference_date: str | None = None,
        sent_at: str | None = None,
        reference_source: str = "user_or_today",
        timezone_name: str = "Asia/Seoul",
        input_source: str = "text",
    ) -> Analysis:
        """Classify `text` and store the drafts as pending work items."""

        result, source, model, warning = self.classify(
            text,
            use_ai=use_ai,
            reference_date=reference_date,
            sent_at=sent_at,
        )
        self._apply_selection_policy(result.work_items)
        if warning and source == "offline":
            self._mark_ai_fallback(result.work_items)
        context: dict[str, object] = {
            "schema_version": 2,
            "prompt_version": "semantic-v4-compact",
            "validation_version": "strict-v1",
            "sent_at": sent_at,
            "reference_date": reference_date,
            "reference_source": reference_source,
            "timezone": timezone_name,
            "input_source": input_source,
            "source": source,
            "model": model,
            "fallback": bool(warning and source == "offline"),
        }
        input_id = self.database.create_input(text, result.work_items, context=context)
        return Analysis(
            input_id=input_id,
            text=text,
            work_items=result.work_items,
            ignored=result.ignored,
            source=source,
            model=model,
            warning=warning,
            context=context,
        )

    @staticmethod
    def _apply_selection_policy(work_items: list[WorkItem]) -> None:
        """Choose safe defaults without silently turning context into work.

        Conditional recruitment/reply candidates and completed or purely
        informational source messages remain visible, but are not selected by
        default. An unresolved Calendar candidate is retained for review and
        cannot reach Outlook until its start time is supplied.
        """

        for item in work_items:
            if item.scope == SCOPE_REFERENCE:
                item.selected = False
                continue
            if item.source_state in {"completed", "informational"}:
                item.selected = False
            if item.applicability != "required":
                item.selected = False
            if item.outlook_target == CALENDAR and not item.start:
                if not any(
                    issue.code == "missing_execution_time"
                    for issue in item.review_issues
                ):
                    item.review_issues.append(
                        ReviewIssue(
                            code="missing_execution_time",
                            field="start",
                            message="수행·참석 시각이 정해지지 않아 Outlook 일정으로 등록할 수 없습니다.",
                            blocking=True,
                        )
                    )
                item.selected = False
            if item.blocking_review_issues:
                item.selected = False

    @staticmethod
    def _mark_ai_fallback(work_items: list[WorkItem]) -> None:
        """Require an explicit click after an AI failure switches to rules."""

        for item in work_items:
            if item.outlook_target is None or item.source_state != "pending":
                continue
            item.selected = False
            if not any(issue.code == "ai_fallback" for issue in item.review_issues):
                item.review_issues.append(
                    ReviewIssue(
                        code="ai_fallback",
                        field="source",
                        message="AI 분석에 실패해 오프라인 규칙 결과를 표시했습니다. 등록 전 내용을 확인하세요.",
                        blocking=False,
                    )
                )

    # --- duplicate guard ---------------------------------------------

    @staticmethod
    def _entry_label(entry: OutlookEntry) -> str:
        label = entry.subject.strip() or "제목 없음"
        if entry.item_type == CALENDAR and entry.start is not None:
            label += f" ({entry.start:%m-%d %H:%M})"
        elif entry.item_type == TODO and entry.due is not None:
            label += f" ({entry.due:%m-%d})"
        return label

    def _conflict_with_entry(
        self, item: WorkItem, entry: OutlookEntry
    ) -> str | None:
        target = item.outlook_target
        if target == CALENDAR and entry.item_type == CALENDAR:
            draft_range = _draft_interval(item)
            entry_range = _entry_interval(entry)
            if draft_range and entry_range and _overlaps(draft_range, entry_range):
                return (
                    "기존 Outlook 일정과 시간이 겹쳐 제외: "
                    f"{self._entry_label(entry)}"
                )
            return None

        if target == TODO and entry.item_type == TODO:
            if _normalise_title(item.title) != _normalise_title(entry.subject):
                return None
            if item.due:
                try:
                    draft_due = parse_outlook_datetime(item.due).date()
                except ValueError:
                    return None
                if entry.due is None or entry.due.date() != draft_due:
                    return None
            return f"기존 Outlook 작업과 같아 제외: {self._entry_label(entry)}"
        return None

    def known_entries(self) -> list[OutlookEntry]:
        """Return what this app believes the Outlook profile currently holds.

        Reads come from the local mirror so that registering one card does not
        pay for a full COM scan.  The mirror is kept current by the background
        sync thread and by write-through after every registration.  Before the
        first sync lands there is nothing to compare against, so that one case
        still reads Outlook directly.
        """

        if self.mirror.has_snapshot():
            self.mirror.refresh_if_stale()
            return self.mirror.entries()
        return self.outlook.read_entries(limit=10000)

    def find_conflicts(self, drafts: list[WorkItem]) -> dict[int, str]:
        """Find Calendar overlaps and exact Todo duplicates before registration.

        Called immediately before presenting the cards and again immediately
        before saving, so an item created between those two moments is still
        caught as long as the mirror has seen it.
        """

        if not any(
            item.selected and not item.saved and item.outlook_target is not None
            for item in drafts
        ):
            return {}

        entries = self.known_entries()
        conflicts: dict[int, str] = {}
        accepted: list[WorkItem] = []

        for index, item in enumerate(drafts):
            if not item.selected or item.saved or item.outlook_target is None:
                continue
            reason: str | None = None
            for entry in entries:
                reason = self._conflict_with_entry(item, entry)
                if reason:
                    break

            # Also remove duplicates produced twice in one message before
            # either one reaches Outlook.
            if reason is None:
                for previous in accepted:
                    if item.outlook_target != previous.outlook_target:
                        continue
                    if item.outlook_target == CALENDAR:
                        left = _draft_interval(item)
                        right = _draft_interval(previous)
                        if left and right and _overlaps(left, right):
                            reason = f"같은 메시지에서 시간이 겹쳐 제외: {previous.title}"
                            break
                    elif (
                        _normalise_title(item.title)
                        == _normalise_title(previous.title)
                        and (
                            not item.due
                            or not previous.due
                            or item.due[:10] == previous.due[:10]
                        )
                    ):
                        reason = f"같은 메시지에서 중복되어 제외: {previous.title}"
                        break

            if reason:
                conflicts[index] = reason
                item.excluded_reason = reason
                item.selected = False
            else:
                accepted.append(item)
        return conflicts

    # --- recording ----------------------------------------------------

    def record(self, drafts: list[WorkItem]) -> RecordOutcome:
        """Write the selected drafts to the current classic Outlook profile.

        Must be called from the thread that owns the Tk loop: Outlook COM is
        apartment-threaded and the adapter is created there.
        """

        outcome = RecordOutcome()
        try:
            conflicts = self.find_conflicts(drafts)
        except (
            AttributeError,
            OSError,
            ValueError,
            OutlookUnavailableError,
            OutlookOperationError,
        ) as exc:
            # The duplicate guard is a safety precondition.  If Outlook
            # cannot be read, proceeding would violate the user's
            # "overlapping items are excluded" rule.
            outcome.blocked = True
            outcome.warning = f"중복 확인을 건너뛰었습니다: {exc}"
            return outcome

        for index, item in enumerate(drafts):
            if index in conflicts:
                outcome.duplicates[index] = conflicts[index]
                item.excluded_reason = conflicts[index]
                item.selected = False
                continue
            if not item.selected or item.saved:
                continue
            if item.outlook_target is None:
                # 참조는 SQLite에만 남기고 Outlook에는 만들지 않는다.
                outcome.skipped += 1
                continue
            try:
                if item.source_state in {"completed", "informational"}:
                    continue
                if item.applicability == "unknown":
                    raise ValueError(
                        "대상이 불명확한 항목은 확인 후 등록해야 합니다."
                    )
                if item.outlook_target == CALENDAR and not item.start:
                    raise ValueError(
                        "Calendar 일정의 수행 시각이 없습니다. 카드의 수정에서 입력하세요."
                    )
                if item.blocking_review_issues:
                    raise ValueError(
                        item.blocking_review_issues[0].message
                        or "확인이 필요한 항목은 먼저 수정해야 합니다."
                    )
                item.validate()
                self.database.update_work_item(item)
                entry_id = self.outlook.save_work_item(item)
                if item.id is None:
                    raise OutlookOperationError(
                        "WorkItem DB ID가 없어 저장 결과를 기록할 수 없습니다."
                    )
                self.database.mark_saved(item.id, entry_id)
                item.saved = True
                item.outlook_entry_id = entry_id
                # 방금 만든 항목을 미러에 바로 반영해 다음 카드의 중복 검사가
                # 다음 주기 동기화를 기다리지 않게 한다.
                self.mirror.write_through(item)
                outcome.saved += 1
                self._add_due_marker(item, outcome)
            except (
                OSError,
                ValueError,
                OutlookUnavailableError,
                OutlookOperationError,
            ) as exc:
                message = str(exc)
                outcome.failed += 1
                outcome.errors[index] = message
                item.last_error = message
                if item.id is not None:
                    try:
                        self.database.mark_failed(item.id, message)
                    except OSError:
                        pass
        return outcome

    def _existing_due_marker(self, item: WorkItem) -> bool:
        """같은 제목·같은 날짜의 마감 표시가 이미 있는가."""

        if not item.due:
            return False
        try:
            due_date = parse_outlook_datetime(item.due).date()
        except ValueError:
            return False
        subject = _normalise_title(due_marker_subject(item.title))
        for entry in self.mirror.entries():
            if entry.item_type != CALENDAR or not _is_due_marker(entry):
                continue
            if entry.start is None or entry.start.date() != due_date:
                continue
            if _normalise_title(entry.subject) == subject:
                return True
        return False

    def _add_due_marker(self, item: WorkItem, outcome: RecordOutcome) -> None:
        """마감 작업을 캘린더에도 표시한다.

        작업은 이미 Outlook에 들어갔으므로, 표시를 만들지 못해도 등록 자체를
        실패로 되돌리지 않습니다. 사유만 경고로 남깁니다.
        """

        if not item.wants_due_marker or item.calendar_entry_id:
            return
        try:
            if self._existing_due_marker(item):
                return
            marker_id = self.outlook.save_due_marker(item)
        except (
            AttributeError,
            OSError,
            ValueError,
            OutlookUnavailableError,
            OutlookOperationError,
        ) as exc:
            note = f"마감 캘린더 표시를 만들지 못했습니다: {exc}"
            outcome.warning = f"{outcome.warning} · {note}" if outcome.warning else note
            return
        item.calendar_entry_id = marker_id
        if item.id is not None:
            self.database.mark_due_marker(item.id, marker_id)
        self.mirror.write_through_due_marker(item)

    def test_outlook(self) -> str:
        return self.outlook.test_connection()
