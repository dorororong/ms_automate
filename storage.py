"""SQLite persistence for inputs and Outlook work items."""

from __future__ import annotations

import sqlite3
from datetime import datetime, timezone
import json
from pathlib import Path

from models import WorkItem


SCHEMA = """
CREATE TABLE IF NOT EXISTS inputs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    original_text TEXT NOT NULL,
    context_json TEXT NOT NULL DEFAULT '{}',
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS work_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    input_id INTEGER NOT NULL REFERENCES inputs(id) ON DELETE CASCADE,
    type TEXT NOT NULL CHECK(type IN ('calendar', 'todo')),
    scope TEXT NOT NULL DEFAULT '업무',
    title TEXT NOT NULL,
    start TEXT,
    end TEXT,
    due TEXT,
    due_time TEXT,
    description TEXT NOT NULL DEFAULT '',
    location TEXT,
    category TEXT,
    all_day INTEGER NOT NULL DEFAULT 0 CHECK(all_day IN (0, 1)),
    reminder_minutes INTEGER,
    reminder TEXT,
    repeat_freq TEXT NOT NULL DEFAULT 'none',
    repeat_detail TEXT NOT NULL DEFAULT '',
    saved INTEGER NOT NULL DEFAULT 0 CHECK(saved IN (0, 1)),
    outlook_entry_id TEXT,
    -- 마감 작업을 캘린더에도 표시할 때 만들어지는 두 번째 Outlook 항목
    calendar_entry_id TEXT,
    last_error TEXT,
    outlook_status TEXT NOT NULL DEFAULT 'pending'
        CHECK(outlook_status IN ('pending', 'saved', 'failed', 'deleted')),
    semantic_json TEXT NOT NULL DEFAULT '{}',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

-- 로컬 Outlook 미러. 중복 검사는 COM 대신 이 표를 읽는다.
CREATE TABLE IF NOT EXISTS outlook_mirror (
    entry_id TEXT PRIMARY KEY,
    item_type TEXT NOT NULL CHECK(item_type IN ('calendar', 'todo')),
    subject TEXT NOT NULL DEFAULT '',
    start TEXT,
    end TEXT,
    due TEXT,
    all_day INTEGER NOT NULL DEFAULT 0 CHECK(all_day IN (0, 1)),
    location TEXT NOT NULL DEFAULT '',
    category TEXT NOT NULL DEFAULT '',
    reminder TEXT,
    reminder_minutes INTEGER,
    complete INTEGER,
    body TEXT NOT NULL DEFAULT '',
    origin TEXT NOT NULL DEFAULT 'outlook' CHECK(origin IN ('outlook', 'app')),
    synced_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_outlook_mirror_type
    ON outlook_mirror(item_type);

CREATE TABLE IF NOT EXISTS sync_state (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
"""

MIRROR_COLUMNS = (
    "entry_id",
    "item_type",
    "subject",
    "start",
    "end",
    "due",
    "all_day",
    "location",
    "category",
    "reminder",
    "reminder_minutes",
    "complete",
    "body",
)


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")


def utc_now_precise_iso() -> str:
    """Microsecond stamp for the mirror.

    A full sync drops rows it did not just see by comparing `synced_at`
    against the moment the Outlook read began, so two writes inside the same
    second have to stay distinguishable.  Every `synced_at` value uses this
    precision so the strings also sort correctly against each other.
    """

    return (
        datetime.now(timezone.utc)
        .isoformat(timespec="microseconds")
        .replace("+00:00", "Z")
    )


class Database:
    def __init__(self, path: str | Path) -> None:
        self.path = Path(path).expanduser()
        self.initialize()

    def _connect(self) -> sqlite3.Connection:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        # 미러 동기화 스레드와 Tk 스레드가 같은 파일에 쓰므로 잠금을 기다린다.
        connection = sqlite3.connect(self.path, timeout=15.0)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        return connection

    def initialize(self) -> None:
        connection = self._connect()
        try:
            connection.executescript(SCHEMA)
            self._migrate_inputs(connection)
            self._migrate_work_items(connection)
            connection.commit()
        finally:
            connection.close()

    @staticmethod
    def _migrate_inputs(connection: sqlite3.Connection) -> None:
        existing = {
            str(row[1])
            for row in connection.execute("PRAGMA table_info(inputs)").fetchall()
        }
        if "context_json" not in existing:
            connection.execute(
                "ALTER TABLE inputs ADD COLUMN context_json TEXT NOT NULL DEFAULT '{}'"
            )

    @staticmethod
    def _migrate_work_items(connection: sqlite3.Connection) -> None:
        """Add fields introduced after the first MVP without losing history."""

        existing = {
            str(row[1])
            for row in connection.execute("PRAGMA table_info(work_items)").fetchall()
        }
        migrations = {
            "location": "ALTER TABLE work_items ADD COLUMN location TEXT",
            "category": "ALTER TABLE work_items ADD COLUMN category TEXT",
            "all_day": (
                "ALTER TABLE work_items ADD COLUMN all_day "
                "INTEGER NOT NULL DEFAULT 0"
            ),
            "reminder_minutes": (
                "ALTER TABLE work_items ADD COLUMN reminder_minutes INTEGER"
            ),
            "reminder": "ALTER TABLE work_items ADD COLUMN reminder TEXT",
            "scope": (
                "ALTER TABLE work_items ADD COLUMN scope TEXT NOT NULL DEFAULT '업무'"
            ),
            "repeat_freq": (
                "ALTER TABLE work_items ADD COLUMN repeat_freq TEXT NOT NULL DEFAULT 'none'"
            ),
            "repeat_detail": (
                "ALTER TABLE work_items ADD COLUMN repeat_detail TEXT NOT NULL DEFAULT ''"
            ),
            "due_time": "ALTER TABLE work_items ADD COLUMN due_time TEXT",
            "calendar_entry_id": (
                "ALTER TABLE work_items ADD COLUMN calendar_entry_id TEXT"
            ),
            "semantic_json": (
                "ALTER TABLE work_items ADD COLUMN semantic_json TEXT NOT NULL DEFAULT '{}'"
            ),
        }
        for column, statement in migrations.items():
            if column not in existing:
                connection.execute(statement)
        if "outlook_status" not in existing:
            connection.execute(
                "ALTER TABLE work_items ADD COLUMN outlook_status "
                "TEXT NOT NULL DEFAULT 'pending'"
            )
            connection.execute(
                "UPDATE work_items SET outlook_status = 'saved' WHERE saved = 1"
            )
            connection.execute(
                "UPDATE work_items SET outlook_status = 'failed' "
                "WHERE saved = 0 AND last_error IS NOT NULL"
            )

    def create_input(
        self,
        original_text: str,
        work_items: list[WorkItem],
        *,
        context: dict[str, object] | None = None,
    ) -> int:
        now = utc_now_iso()
        connection = self._connect()
        try:
            cursor = connection.execute(
                "INSERT INTO inputs(original_text, context_json, created_at) VALUES (?, ?, ?)",
                (
                    original_text,
                    json.dumps(context or {}, ensure_ascii=False),
                    now,
                ),
            )
            input_id = int(cursor.lastrowid)
            for work_item in work_items:
                work_item.validate()
                item_cursor = connection.execute(
                    """
                    INSERT INTO work_items(
                        input_id, type, scope, title, start, end, due, due_time, description,
                        location, category, all_day, reminder_minutes, reminder,
                        repeat_freq, repeat_detail,
                        saved, outlook_entry_id, last_error, outlook_status,
                        semantic_json, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, NULL, NULL, 'pending', ?, ?, ?)
                    """,
                    (
                        input_id,
                        work_item.type,
                        work_item.scope,
                        work_item.title,
                        work_item.start,
                        work_item.end,
                        work_item.due,
                        work_item.due_time,
                        work_item.description,
                        work_item.location,
                        work_item.category,
                        int(work_item.all_day),
                        work_item.reminder_minutes,
                        work_item.reminder,
                        work_item.repeat_freq,
                        work_item.repeat_detail,
                        json.dumps(work_item.semantic_payload(), ensure_ascii=False),
                        now,
                        now,
                    ),
                )
                work_item.id = int(item_cursor.lastrowid)
                work_item.input_id = input_id
            connection.commit()
            return input_id
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()

    def update_work_item(self, work_item: WorkItem) -> None:
        if work_item.id is None:
            raise ValueError("DB에 저장된 WorkItem ID가 없습니다.")
        work_item.validate()
        connection = self._connect()
        try:
            connection.execute(
                """
                UPDATE work_items
                   SET type = ?, scope = ?, title = ?, start = ?, end = ?, due = ?,
                       description = ?, location = ?, category = ?, all_day = ?,
                       due_time = ?, semantic_json = ?,
                       reminder_minutes = ?, reminder = ?,
                       repeat_freq = ?, repeat_detail = ?,
                       updated_at = ?, last_error = ?
                 WHERE id = ?
                """,
                (
                    work_item.type,
                    work_item.scope,
                    work_item.title,
                    work_item.start,
                    work_item.end,
                    work_item.due,
                    work_item.description,
                    work_item.location,
                    work_item.category,
                    int(work_item.all_day),
                    work_item.due_time,
                    json.dumps(work_item.semantic_payload(), ensure_ascii=False),
                    work_item.reminder_minutes,
                    work_item.reminder,
                    work_item.repeat_freq,
                    work_item.repeat_detail,
                    utc_now_iso(),
                    work_item.last_error,
                    work_item.id,
                ),
            )
            connection.commit()
        finally:
            connection.close()

    def mark_saved(self, work_item_id: int, outlook_entry_id: str | None) -> None:
        connection = self._connect()
        try:
            connection.execute(
                """
                UPDATE work_items
                   SET saved = 1, outlook_entry_id = ?, last_error = NULL,
                       outlook_status = 'saved', updated_at = ?
                 WHERE id = ?
                """,
                (outlook_entry_id, utc_now_iso(), work_item_id),
            )
            connection.commit()
        finally:
            connection.close()

    def mark_due_marker(self, work_item_id: int, entry_id: str | None) -> None:
        """마감 캘린더 표시의 EntryID를 남긴다. 작업 자체의 등록과는 별개다."""

        connection = self._connect()
        try:
            connection.execute(
                "UPDATE work_items SET calendar_entry_id = ?, updated_at = ? WHERE id = ?",
                (entry_id, utc_now_iso(), work_item_id),
            )
            connection.commit()
        finally:
            connection.close()

    def mark_failed(self, work_item_id: int, error: str) -> None:
        connection = self._connect()
        try:
            connection.execute(
                "UPDATE work_items SET saved = 0, last_error = ?, "
                "outlook_status = 'failed', updated_at = ? WHERE id = ?",
                (error[:1000], utc_now_iso(), work_item_id),
            )
            connection.commit()
        finally:
            connection.close()

    def mark_deleted(self, work_item_id: int) -> None:
        """Mark an app-managed Outlook item as deleted without losing its audit row."""

        connection = self._connect()
        try:
            connection.execute(
                "UPDATE work_items SET saved = 0, last_error = NULL, "
                "outlook_status = 'deleted', updated_at = ? WHERE id = ?",
                (utc_now_iso(), work_item_id),
            )
            connection.commit()
        finally:
            connection.close()

    def get_work_items(self, input_id: int) -> list[WorkItem]:
        connection = self._connect()
        try:
            rows = connection.execute(
                "SELECT * FROM work_items WHERE input_id = ? ORDER BY id", (input_id,)
            ).fetchall()
            return [WorkItem.from_row(row) for row in rows]
        finally:
            connection.close()

    def get_saved_work_items(self) -> list[WorkItem]:
        """Return Outlook items created by this app and still marked as saved."""

        connection = self._connect()
        try:
            rows = connection.execute(
                "SELECT * FROM work_items "
                "WHERE saved = 1 AND outlook_entry_id IS NOT NULL "
                "ORDER BY id"
            ).fetchall()
            return [WorkItem.from_row(row) for row in rows]
        finally:
            connection.close()

    # --- Outlook 미러 --------------------------------------------------

    def _upsert_mirror_rows(
        self,
        connection: sqlite3.Connection,
        rows: list[dict[str, object]],
        *,
        origin: str,
        synced_at: str,
    ) -> None:
        columns = ", ".join(MIRROR_COLUMNS)
        placeholders = ", ".join("?" for _ in MIRROR_COLUMNS)
        assignments = ", ".join(
            f"{name} = excluded.{name}" for name in MIRROR_COLUMNS[1:]
        )
        connection.executemany(
            f"""
            INSERT INTO outlook_mirror ({columns}, origin, synced_at)
            VALUES ({placeholders}, ?, ?)
            ON CONFLICT(entry_id) DO UPDATE SET
                {assignments}, origin = excluded.origin, synced_at = excluded.synced_at
            """,
            [
                tuple(row.get(name) for name in MIRROR_COLUMNS) + (origin, synced_at)
                for row in rows
            ],
        )

    def sync_outlook_mirror(
        self, rows: list[dict[str, object]], *, started_at: str
    ) -> int:
        """Replace the mirror with one full Outlook read.

        Rows written by this app while the read was running carry a newer
        `synced_at` than `started_at`, so they survive the cleanup that drops
        items which disappeared from Outlook.
        """

        synced_at = utc_now_precise_iso()
        connection = self._connect()
        try:
            self._upsert_mirror_rows(
                connection, rows, origin="outlook", synced_at=synced_at
            )
            connection.execute(
                "DELETE FROM outlook_mirror WHERE synced_at < ?", (started_at,)
            )
            connection.execute(
                "INSERT INTO sync_state (key, value, updated_at) VALUES (?, ?, ?) "
                "ON CONFLICT(key) DO UPDATE SET value = excluded.value, "
                "updated_at = excluded.updated_at",
                ("outlook_mirror_synced_at", synced_at, synced_at),
            )
            connection.commit()
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()
        return len(rows)

    def upsert_outlook_entry(self, row: dict[str, object]) -> None:
        """Write through one item this app just created or changed in Outlook."""

        connection = self._connect()
        try:
            self._upsert_mirror_rows(
                connection, [row], origin="app", synced_at=utc_now_precise_iso()
            )
            connection.commit()
        finally:
            connection.close()

    def delete_outlook_entry(self, entry_id: str) -> None:
        connection = self._connect()
        try:
            connection.execute(
                "DELETE FROM outlook_mirror WHERE entry_id = ?", (entry_id,)
            )
            connection.commit()
        finally:
            connection.close()

    def read_outlook_mirror(self) -> list[sqlite3.Row]:
        connection = self._connect()
        try:
            return connection.execute(
                "SELECT * FROM outlook_mirror ORDER BY start, due, subject"
            ).fetchall()
        finally:
            connection.close()

    def count_outlook_mirror(self) -> int:
        connection = self._connect()
        try:
            row = connection.execute(
                "SELECT COUNT(*) AS count FROM outlook_mirror"
            ).fetchone()
            return int(row["count"])
        finally:
            connection.close()

    def get_sync_state(self, key: str) -> str | None:
        connection = self._connect()
        try:
            row = connection.execute(
                "SELECT value FROM sync_state WHERE key = ?", (key,)
            ).fetchone()
            return None if row is None else str(row["value"])
        finally:
            connection.close()

    def count_inputs(self) -> int:
        connection = self._connect()
        try:
            row = connection.execute("SELECT COUNT(*) AS count FROM inputs").fetchone()
            return int(row["count"])
        finally:
            connection.close()
