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
    last_error TEXT,
    outlook_status TEXT NOT NULL DEFAULT 'pending'
        CHECK(outlook_status IN ('pending', 'saved', 'failed', 'deleted')),
    semantic_json TEXT NOT NULL DEFAULT '{}',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
"""


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")


class Database:
    def __init__(self, path: str | Path) -> None:
        self.path = Path(path).expanduser()
        self.initialize()

    def _connect(self) -> sqlite3.Connection:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        connection = sqlite3.connect(self.path)
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

    def count_inputs(self) -> int:
        connection = self._connect()
        try:
            row = connection.execute("SELECT COUNT(*) AS count FROM inputs").fetchone()
            return int(row["count"])
        finally:
            connection.close()
