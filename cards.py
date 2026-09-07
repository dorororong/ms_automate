"""Compact editable card widget for one WorkItem draft."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Callable

import theme
from models import (
    CALENDAR,
    ACTION_NONE,
    ACTION_TODO,
    ACTION_WITH_DATE,
    ACTION_WITH_DUE,
    REPEAT_FREQS,
    SCOPE_REFERENCE,
    SCOPES,
    ReviewIssue,
    WorkItem,
)

ACTION_BADGE = {
    ACTION_TODO: ("할 일", "Outlook 작업 · 마감 없음"),
    ACTION_WITH_DATE: ("일정", "Outlook 캘린더 · 알림"),
    ACTION_WITH_DUE: ("마감", "Outlook 작업 · 마감일 + 내역"),
    ACTION_NONE: ("참조", "등록하지 않음 · 로컬 기록만"),
}


def display_datetime(value: str | None) -> str:
    """저장형 일시를 화면용 문자열로 바꾼다."""

    if not value:
        return ""
    text = str(value).replace("T", " ").strip()
    if text.count(":") == 2:  # 초는 보여줄 필요가 없다.
        text = text.rsplit(":", 1)[0]
    return text


class WorkItemCard:
    """한 건의 결과를 짧게 보여주고 필요할 때만 상세 편집을 연다."""

    def __init__(
        self,
        parent: tk.Misc,
        item: WorkItem,
        index: int,
        *,
        on_change: Callable[[], None] | None = None,
    ) -> None:
        self.item = item
        self.index = index
        self.on_change = on_change
        accent = theme.SCOPE_COLORS.get(item.scope, theme.MUTED)
        self.frame, body = theme.card(parent, accent)
        self.strip = self.frame.winfo_children()[0]

        self.selected_var = tk.BooleanVar(
            value=item.selected and not item.saved and not item.excluded_reason
        )
        self.scope_var = tk.StringVar(value=item.scope)
        self.title_var = tk.StringVar(value=item.title)
        self.at_var = tk.StringVar(value=display_datetime(item.start))
        self.due_var = tk.StringVar(value=display_datetime(item.due))
        self.due_time_var = tk.StringVar(value=item.due_time or "")
        self.location_var = tk.StringVar(value=item.location or "")
        self.all_day_var = tk.BooleanVar(value=item.all_day)
        self.reminder_var = tk.StringVar(
            value=str(item.reminder_minutes) if item.reminder_minutes is not None else ""
        )
        self.repeat_freq_var = tk.StringVar(value=item.repeat_freq)
        self.repeat_detail_var = tk.StringVar(value=item.repeat_detail)
        self.compact_meta_var = tk.StringVar(value="")
        self._details_visible = False
        self.selected_var.trace_add("write", lambda *_args: self._notify_change())

        body.columnconfigure(0, weight=1)

        # --- 항상 보이는 요약 ------------------------------------------
        top = ttk.Frame(body, style="Card.TFrame")
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)
        self.select_button = theme.checkbox(top, "등록", self.selected_var)
        self.select_button.grid(row=0, column=0, padx=(0, 8))
        self.title_label = ttk.Label(
            top, textvariable=self.title_var, style="CardHead.TLabel", anchor="w"
        )
        self.title_label.grid(row=0, column=1, sticky="ew")
        self.toggle_button = ttk.Button(top, text="수정", width=7, command=self._toggle_details)
        self.toggle_button.grid(row=0, column=2, padx=(8, 0))

        meta = ttk.Frame(body, style="Card.TFrame")
        meta.grid(row=1, column=0, sticky="ew", pady=(5, 0))
        meta.columnconfigure(1, weight=1)
        self.badge = tk.Label(meta, text="", font=theme.F_SMALL, padx=7, pady=1)
        self.badge.grid(row=0, column=0, sticky="w")
        ttk.Label(
            meta, textvariable=self.compact_meta_var, style="CardMuted.TLabel", anchor="w"
        ).grid(row=0, column=1, sticky="ew", padx=(8, 0))
        self.state_label = ttk.Label(meta, text="", style="CardMuted.TLabel", anchor="e")
        self.state_label.grid(row=0, column=2, sticky="e", padx=(8, 0))

        # --- 필요할 때만 여는 상세 편집 -------------------------------
        self.details = ttk.Frame(body, style="Card.TFrame", padding=(0, 8, 0, 0))
        self.details.grid(row=2, column=0, sticky="ew")
        self.details.columnconfigure(1, weight=1)

        ttk.Label(self.details, text="제목", style="Card.TLabel", width=5).grid(
            row=0, column=0, sticky="w", pady=2
        )
        title_entry = ttk.Entry(self.details, textvariable=self.title_var, font=theme.F_BODY)
        title_entry.grid(row=0, column=1, columnspan=3, sticky="ew", pady=2)
        title_entry.bind("<KeyRelease>", self._refresh)
        self.title_entry = title_entry

        ttk.Label(self.details, text="분류", style="Card.TLabel", width=5).grid(
            row=1, column=0, sticky="w", pady=2
        )
        scope_box = ttk.Frame(self.details, style="Card.TFrame")
        scope_box.grid(row=1, column=1, columnspan=3, sticky="w", pady=2)
        for scope in SCOPES:
            ttk.Radiobutton(
                scope_box,
                text=scope,
                value=scope,
                variable=self.scope_var,
                command=self._refresh,
                style="Card.TRadiobutton",
            ).pack(side="left", padx=(0, 8))
        self.badge_hint = ttk.Label(scope_box, text="", style="CardMuted.TLabel")
        self.badge_hint.pack(side="left", padx=(4, 0))

        ttk.Label(self.details, text="일시", style="Card.TLabel", width=5).grid(
            row=2, column=0, sticky="w", pady=2
        )
        at_entry = ttk.Entry(self.details, textvariable=self.at_var, width=19)
        at_entry.grid(row=2, column=1, sticky="w", pady=2)
        at_entry.bind("<KeyRelease>", self._refresh)
        ttk.Label(self.details, text="마감", style="Card.TLabel", width=5).grid(
            row=2, column=2, sticky="w", padx=(12, 0), pady=2
        )
        due_entry = ttk.Entry(self.details, textvariable=self.due_var, width=14)
        due_entry.grid(row=2, column=3, sticky="w", pady=2)
        due_entry.bind("<KeyRelease>", self._refresh)

        ttk.Label(self.details, text="마감 시각", style="Card.TLabel", width=7).grid(
            row=3, column=0, sticky="w", pady=2
        )
        ttk.Entry(self.details, textvariable=self.due_time_var, width=8).grid(
            row=3, column=1, sticky="w", pady=2
        )
        ttk.Checkbutton(
            self.details,
            text="종일",
            variable=self.all_day_var,
            command=self._refresh,
        ).grid(row=3, column=2, columnspan=2, sticky="w", padx=(12, 0), pady=2)

        ttk.Label(self.details, text="장소", style="Card.TLabel", width=5).grid(
            row=4, column=0, sticky="w", pady=2
        )
        ttk.Entry(self.details, textvariable=self.location_var, width=19).grid(
            row=4, column=1, sticky="w", pady=2
        )
        ttk.Label(self.details, text="알림", style="Card.TLabel", width=5).grid(
            row=4, column=2, sticky="w", padx=(12, 0), pady=2
        )
        ttk.Entry(self.details, textvariable=self.reminder_var, width=6).grid(
            row=4, column=3, sticky="w", pady=2
        )

        ttk.Label(self.details, text="반복", style="Card.TLabel", width=5).grid(
            row=5, column=0, sticky="w", pady=2
        )
        ttk.Combobox(
            self.details,
            textvariable=self.repeat_freq_var,
            values=REPEAT_FREQS,
            width=10,
            state="readonly",
        ).grid(row=5, column=1, sticky="w", pady=2)
        ttk.Entry(self.details, textvariable=self.repeat_detail_var, width=13).grid(
            row=5, column=2, columnspan=2, sticky="w", padx=(12, 0), pady=2
        )

        ttk.Label(self.details, text="내역", style="Card.TLabel", width=5).grid(
            row=6, column=0, sticky="nw", pady=(5, 2)
        )
        description_wrap = tk.Frame(self.details, bg=theme.BORDER)
        description_wrap.grid(row=6, column=1, columnspan=3, sticky="ew", pady=(5, 2))
        description_wrap.columnconfigure(0, weight=1)
        self.description_text = tk.Text(
            description_wrap,
            height=2,
            wrap="word",
            font=theme.F_BODY,
            relief="flat",
            highlightthickness=0,
            padx=6,
            pady=4,
            background=theme.SURFACE,
            foreground=theme.TEXT,
        )
        self.description_text.grid(row=0, column=0, sticky="ew", padx=1, pady=1)
        self.description_text.insert("1.0", item.description)
        self.details.grid_remove()

        self._excluded_signature: tuple[str, ...] | None = (
            self._current_signature() if item.excluded_reason else None
        )
        self._refresh()
        if item.saved:
            self.select_button.configure(state="disabled")

    def _notify_change(self) -> None:
        if self.on_change is not None:
            self.on_change()

    # --- 접기/펼치기 ----------------------------------------------------

    def _toggle_details(self) -> None:
        self._details_visible = not self._details_visible
        if self._details_visible:
            self.details.grid()
            self.toggle_button.configure(text="접기")
            self.title_entry.focus_set()
        else:
            self.details.grid_remove()
            self.toggle_button.configure(text="수정")

    # --- 파생 라벨 ------------------------------------------------------

    def _refresh(self, _event: tk.Event | None = None) -> None:
        if self.item.type == CALENDAR:
            has_start = bool(self.at_var.get().strip())
            if not has_start and not any(
                issue.code == "missing_execution_time"
                for issue in self.item.review_issues
            ):
                self.item.review_issues.append(
                    ReviewIssue(
                        code="missing_execution_time",
                        field="start",
                        message="수행·참석 시각이 없어 등록 전에 확인이 필요합니다.",
                        blocking=True,
                    )
                )
            for issue in self.item.review_issues:
                if issue.code == "missing_execution_time":
                    issue.resolved = has_start
        has_due = bool(self.due_var.get().strip())
        for issue in self.item.review_issues:
            if issue.code == "ambiguous_deadline_date":
                issue.resolved = has_due
        if self.item.excluded_reason and self._current_signature() != self._excluded_signature:
            # 제목·일시·분류를 고치는 것은 중복을 해결하려는 명시적 수정이다.
            self.item.excluded_reason = None
            self._excluded_signature = None

        scope = self.scope_var.get()
        preview = WorkItem(
            type=CALENDAR if self.item.type == CALENDAR and not self.due_var.get().strip() else "todo",
            scope=scope,
            title=self.title_var.get() or "-",
            start=self.at_var.get().strip() or None,
            due=self.due_var.get().strip() or None,
            due_time=self.due_time_var.get().strip() or None,
        )
        label, hint = ACTION_BADGE[preview.action]
        accent = theme.SCOPE_COLORS.get(scope, theme.MUTED)
        self.badge.configure(text=label, bg=accent, fg="#ffffff")
        self.badge_hint.configure(text=hint)
        self.strip.configure(bg=accent)
        self.compact_meta_var.set(self._compact_summary(preview))

        if self.item.excluded_reason:
            self.selected_var.set(False)
            self.select_button.configure(state="disabled")
            self.state_label.configure(
                text=f"중복 제외 · {self.item.excluded_reason}", foreground=theme.WARN
            )
        elif scope == SCOPE_REFERENCE:
            self.selected_var.set(False)
            self.select_button.configure(state="disabled")
            self.state_label.configure(text="참조 · 등록 안 함", foreground=theme.MUTED)
        elif self.item.saved:
            self.selected_var.set(False)
            self.select_button.configure(state="disabled")
            self.state_label.configure(text="등록 완료", foreground=theme.OK)
        elif self.item.blocking_review_issues:
            self.selected_var.set(False)
            self.select_button.configure(state="disabled")
            self.state_label.configure(text="확인 필요", foreground=theme.WARN)
        elif self.item.applicability == "unknown":
            self.selected_var.set(False)
            self.select_button.configure(state="disabled")
            self.state_label.configure(text="대상 확인 필요", foreground=theme.WARN)
        elif self.item.source_state in {"completed", "informational"}:
            self.selected_var.set(False)
            self.select_button.configure(state="disabled")
            self.state_label.configure(text="완료/안내 · 등록 안 함", foreground=theme.MUTED)
        elif self.item.applicability == "conditional":
            self.select_button.configure(state="normal")
            self.state_label.configure(text="조건부 · 선택 시 등록", foreground=theme.WARN)
        elif self.item.last_error:
            self.select_button.configure(state="normal")
            self.state_label.configure(text="등록 실패 · 다시 시도", foreground=theme.FAIL)
        else:
            self.select_button.configure(state="normal")
            self.state_label.configure(text="등록 대기", foreground=theme.MUTED)

    def _compact_summary(self, preview: WorkItem) -> str:
        parts: list[str] = []
        at = self.at_var.get().strip()
        due = self.due_var.get().strip()
        if preview.action == ACTION_WITH_DATE:
            parts.append(display_datetime(at) if at else "날짜 확인 필요")
        elif preview.action == ACTION_WITH_DUE:
            due_label = display_datetime(due) if due else "마감 확인 필요"
            if self.due_time_var.get().strip():
                due_label += f" {self.due_time_var.get().strip()}"
            parts.append(f"마감 {due_label}")
        elif at:
            parts.append(display_datetime(at))
        elif due:
            parts.append(f"마감 {display_datetime(due)}")
        if self.due_time_var.get().strip() and not due:
            parts.append(f"마감 시각 {self.due_time_var.get().strip()} · 날짜 확인")
        if self.item.applicability == "conditional":
            parts.append("조건부")
        for context in self.item.temporal_context:
            if context.role not in {"event_context", "external_deadline", "constraint"}:
                continue
            value = " ".join(
                part for part in (context.date, context.time) if part
            ) or context.raw_text
            if value:
                label = "행사" if context.role == "event_context" else "관련"
                parts.append(f"{label} {value}")
        if self.location_var.get().strip():
            parts.append(self.location_var.get().strip())
        return " · ".join(parts) or "날짜 없음"

    def _current_signature(self) -> tuple[str, ...]:
        return (
            self.title_var.get().strip(),
            self.at_var.get().strip(),
            self.due_var.get().strip(),
            self.scope_var.get().strip(),
        )

    # --- 값 꺼내기 ------------------------------------------------------

    def to_work_item(self) -> WorkItem:
        review_issues = list(self.item.review_issues)
        if self.at_var.get().strip():
            for issue in review_issues:
                if issue.code == "missing_execution_time":
                    issue.resolved = True
        if self.due_var.get().strip():
            for issue in review_issues:
                if issue.code == "ambiguous_deadline_date":
                    issue.resolved = True
        item = WorkItem(
            id=self.item.id,
            input_id=self.item.input_id,
            type="todo",
            scope=self.scope_var.get().strip(),
            title=self.title_var.get().strip(),
            start=self.at_var.get().strip() or None,
            end=self.item.end,
            due=self.due_var.get().strip() or None,
            due_time=self.due_time_var.get().strip() or None,
            description=self.description_text.get("1.0", "end-1c").strip(),
            location=self.location_var.get().strip() or None,
            category=self.item.category,
            all_day=self.all_day_var.get(),
            reminder_minutes=self._parse_reminder_minutes(),
            reminder=self.item.reminder,
            repeat_freq=self.repeat_freq_var.get().strip() or "none",
            repeat_detail=self.repeat_detail_var.get().strip(),
            selected=self.selected_var.get(),
            saved=self.item.saved,
            outlook_entry_id=self.item.outlook_entry_id,
            last_error=None,
            outlook_status=self.item.outlook_status,
            excluded_reason=self.item.excluded_reason,
            intent=self.item.intent,
            applicability=self.item.applicability,
            condition=self.item.condition,
            source_state=self.item.source_state,
            temporal_context=list(self.item.temporal_context),
            checklist=list(self.item.checklist),
            evidence=list(self.item.evidence),
            review_issues=review_issues,
            group_id=self.item.group_id,
        )
        if self.item.type == CALENDAR and not item.due:
            item.type = CALENDAR
        else:
            item.type = item.outlook_target or "todo"
        return item

    def _parse_reminder_minutes(self) -> int | None:
        value = self.reminder_var.get().strip()
        if not value:
            return None
        try:
            return int(value)
        except ValueError as exc:
            raise ValueError("알림은 분 단위 숫자로 입력하세요.") from exc

    # --- 결과 표시 ------------------------------------------------------

    def mark_saved(self, entry_id: str | None) -> None:
        self.item.saved = True
        self.item.outlook_entry_id = entry_id
        self.selected_var.set(False)
        self.select_button.configure(state="disabled")
        self.state_label.configure(text="등록 완료", foreground=theme.OK)

    def mark_excluded(self, reason: str) -> None:
        self.item.excluded_reason = reason
        self._excluded_signature = self._current_signature()
        self.selected_var.set(False)
        self.select_button.configure(state="disabled")
        self.state_label.configure(text=f"중복 제외 · {reason}", foreground=theme.WARN)

    def mark_failed(self, error: str) -> None:
        self.item.last_error = error
        self.state_label.configure(text=f"등록 실패 · {error}", foreground=theme.FAIL)
