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

# 제목이 아무리 길어도 카드가 화면을 다 먹지 않도록 줄 수를 제한한다.
MAX_TITLE_LINES = 4

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
        on_register: Callable[["WorkItemCard"], None] | None = None,
        on_dismiss: Callable[["WorkItemCard"], None] | None = None,
    ) -> None:
        self.item = item
        self.index = index
        self.on_change = on_change
        self.on_register = on_register
        self.on_dismiss = on_dismiss
        self.dismissed = False
        self._title_fit_pending = False
        self._fitting_title = False
        accent = theme.SCOPE_COLORS.get(item.scope, theme.MUTED)
        self.frame, body = theme.card(parent, accent)
        self.strip = self.frame.winfo_children()[0]

        self.selected_var = tk.BooleanVar(
            value=item.selected and not item.saved and not item.excluded_reason
        )
        self.scope_var = tk.StringVar(value=item.scope)
        # 제목은 Text 위젯이 원본이다. StringVar 로 복사해 두면 붙여넣기·되돌리기
        # 처럼 KeyRelease 가 없는 편집에서 값이 어긋난다.
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
        self.show_due_var = tk.BooleanVar(value=item.show_due_on_calendar)
        self.compact_meta_var = tk.StringVar(value="")
        self._details_visible = False
        self.selected_var.trace_add("write", lambda *_args: self._notify_change())

        body.columnconfigure(0, weight=1)

        # --- 항상 보이는 요약 ------------------------------------------
        # 요약 제목은 길면 한 줄에 안 들어간다. 카드 전체 폭을 쓰고 줄바꿈해서
        # 상세를 열지 않아도 제목 전체가 보이게 한다.
        title_wrap = tk.Frame(body, bg=theme.BORDER)
        title_wrap.grid(row=0, column=0, sticky="ew")
        title_wrap.columnconfigure(0, weight=1)
        self.title_text = tk.Text(
            title_wrap,
            height=1,
            # 요청 폭을 1로 두어야 카드 폭에 맞춰 줄바꿈된다. 기본 80자를
            # 그대로 두면 위젯이 카드보다 넓어져 잘려 보이기만 한다.
            width=1,
            wrap="word",
            font=theme.F_HEAD,
            relief="flat",
            highlightthickness=0,
            padx=6,
            pady=4,
            background=theme.SURFACE,
            foreground=theme.TEXT,
            insertbackground=theme.TEXT,
        )
        self.title_text.grid(row=0, column=0, sticky="ew", padx=1, pady=1)
        self.title_text.insert("1.0", item.title)
        self.title_text.bind("<KeyRelease>", self._on_title_edited)
        # Configure 안에서 바로 높이를 재면 배치가 재귀한다. 유휴 시점으로 미룬다.
        self.title_text.bind("<Configure>", self._schedule_title_fit)
        # 제목 칸에서 줄바꿈은 의미가 없다. 엔터로 카드가 커지지 않게 막는다.
        self.title_text.bind("<Return>", lambda _e: "break")
        self.title_entry = self.title_text

        meta = ttk.Frame(body, style="Card.TFrame")
        meta.grid(row=1, column=0, sticky="ew", pady=(5, 0))
        meta.columnconfigure(1, weight=1)
        # 배지가 이미 일정/마감/할 일을 말하므로 같은 말을 옆에 또 쓰지 않는다.
        self.badge = tk.Label(meta, text="", font=theme.F_SMALL, padx=7, pady=1)
        self.badge.grid(row=0, column=0, sticky="w")
        self.quick_date_entry = ttk.Entry(meta, width=17)
        self.quick_date_entry.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        self.quick_date_entry.bind("<KeyRelease>", self._refresh)
        self.quick_time_label = ttk.Label(meta, text="시각", style="CardMuted.TLabel")
        self.quick_time_label.grid(row=0, column=2, sticky="w", padx=(6, 0))
        self.quick_time_entry = ttk.Entry(meta, textvariable=self.due_time_var, width=7)
        self.quick_time_entry.grid(row=0, column=3, sticky="w", padx=(4, 0))
        self.quick_time_entry.bind("<KeyRelease>", self._refresh)
        actions = ttk.Frame(body, style="Card.TFrame")
        actions.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        actions.columnconfigure(0, weight=1)
        self.compact_meta_label = ttk.Label(
            actions, textvariable=self.compact_meta_var, style="CardMuted.TLabel", anchor="w"
        )
        self.compact_meta_label.grid(row=0, column=0, sticky="ew")
        self.toggle_button = ttk.Button(
            actions, text="상세", width=5, command=self._toggle_details
        )
        self.toggle_button.grid(row=0, column=1, padx=(8, 0))
        self.dismiss_button = ttk.Button(
            actions, text="미등록", width=7, command=self._on_dismiss_clicked
        )
        self.dismiss_button.grid(row=0, column=2, padx=(8, 0))
        # 등록은 되돌리기 어려우므로 미등록과 확실히 떨어뜨려 둔다.
        self.register_button = ttk.Button(
            actions, text="등록", width=6, style="Accent.TButton",
            command=self._on_register_clicked,
        )
        self.register_button.grid(row=0, column=3, padx=(26, 0))
        self.state_label = ttk.Label(meta, text="", style="CardMuted.TLabel", anchor="e")
        self.state_label.grid(row=0, column=4, sticky="e", padx=(8, 0))

        # --- 필요할 때만 여는 상세 편집 -------------------------------
        self.details = ttk.Frame(body, style="Card.TFrame", padding=(0, 8, 0, 0))
        self.details.grid(row=3, column=0, sticky="ew")
        self.details.columnconfigure(1, weight=1)

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
        theme.checkbox(
            self.details, "종일", self.all_day_var, command=self._refresh
        ).grid(row=3, column=2, columnspan=2, sticky="w", padx=(12, 0), pady=2)

        ttk.Label(self.details, text="장소", style="Card.TLabel", width=5).grid(
            row=4, column=0, sticky="w", pady=2
        )
        self.location_entry = ttk.Entry(
            self.details, textvariable=self.location_var, width=19
        )
        self.location_entry.grid(row=4, column=1, sticky="w", pady=2)
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

        # 마감 작업은 Outlook 캘린더에 안 보이므로 마감일 종일 표시를 함께 만든다.
        self.due_marker_check = theme.checkbox(
            self.details,
            "마감일을 캘린더에도 표시",
            self.show_due_var,
            command=self._refresh,
        )
        self.due_marker_check.grid(
            row=6, column=0, columnspan=4, sticky="w", pady=(4, 2)
        )

        ttk.Label(self.details, text="내역", style="Card.TLabel", width=5).grid(
            row=7, column=0, sticky="nw", pady=(5, 2)
        )
        description_wrap = tk.Frame(self.details, bg=theme.BORDER)
        description_wrap.grid(row=7, column=1, columnspan=3, sticky="ew", pady=(5, 2))
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
        self._fit_title_height()
        if item.saved:
            self._set_register_enabled(False)

    def title(self) -> str:
        """항상 제목 위젯에서 직접 읽는다."""

        try:
            return self.title_text.get("1.0", "end-1c").strip()
        except tk.TclError:
            return self.item.title

    def _on_title_edited(self, event: tk.Event | None = None) -> None:
        self._fit_title_height()
        self._refresh(event)

    def _schedule_title_fit(self, _event: tk.Event | None = None) -> None:
        if self._title_fit_pending:
            return
        self._title_fit_pending = True
        try:
            self.title_text.after_idle(self._run_title_fit)
        except tk.TclError:
            self._title_fit_pending = False

    def _run_title_fit(self) -> None:
        self._title_fit_pending = False
        self._fit_title_height()

    def _fit_title_height(self) -> None:
        """제목이 길면 줄을 늘려 잘리지 않게 한다.

        `count -displaylines` 는 위젯이 실제로 보여줄 수 있는 줄만 세기 때문에
        높이가 1이면 언제나 1을 돌려준다. 그래서 마지막 글자가 보일 때까지
        높이를 한 줄씩 올려 본다.
        """

        # 아래 update_idletasks 가 배치를 다시 돌리면서 이 함수를 또 부른다.
        # 재진입을 막지 않으면 무한 재귀가 된다.
        if self._fitting_title:
            return
        self._fitting_title = True
        try:
            last = "end-2c" if self.title_text.get("1.0", "end-1c") else "1.0"
            for wanted in range(1, MAX_TITLE_LINES + 1):
                if int(self.title_text.cget("height")) != wanted:
                    self.title_text.configure(height=wanted)
                self.title_text.update_idletasks()
                if self.title_text.bbox(last) is not None:
                    return
        except tk.TclError:  # 카드가 이미 없어졌다.
            return
        finally:
            self._fitting_title = False

    def _notify_change(self) -> None:
        if self.on_change is not None:
            self.on_change()

    # --- 카드에서 바로 처리 ---------------------------------------------

    def _set_register_enabled(self, enabled: bool) -> None:
        """등록 버튼 하나로 이 카드를 지금 저장할 수 있는지 보여준다."""

        allowed = enabled and not self.dismissed and not self.item.saved
        self.register_button.configure(state="normal" if allowed else "disabled")
        if self.item.saved:
            self.dismiss_button.configure(state="disabled", text="미등록")
        else:
            self.dismiss_button.configure(
                state="normal", text="되돌리기" if self.dismissed else "미등록"
            )

    def _on_register_clicked(self) -> None:
        if self.on_register is not None:
            self.on_register(self)

    def _on_dismiss_clicked(self) -> None:
        """미등록은 이 카드만 접어 두는 결정이며 되돌릴 수 있다."""

        if self.item.saved:
            return
        self.dismissed = not self.dismissed
        if self.dismissed:
            self.selected_var.set(False)
            if self._details_visible:
                self._toggle_details()
        self._refresh()
        if self.on_dismiss is not None:
            self.on_dismiss(self)

    def is_pending(self) -> bool:
        """등록도 미등록도 아직 결정되지 않은 카드.

        중복으로 제외된 카드는 결정이 끝난 것으로 본다. 제목이나 시간을 고치면
        `_refresh`가 제외 사유를 지우므로 다시 대기 상태가 된다.
        """

        return (
            not self.item.saved
            and not self.dismissed
            and not self.item.excluded_reason
        )

    def refresh_state(self) -> None:
        """편집값을 기준으로 버튼과 상태 라벨을 다시 계산한다."""

        self._refresh()

    # --- 접기/펼치기 ----------------------------------------------------

    def _toggle_details(self) -> None:
        self._details_visible = not self._details_visible
        if self._details_visible:
            self.details.grid()
            self.toggle_button.configure(text="접기")
            # 제목은 카드 위에 그대로 있으므로 상세를 열면 첫 상세 칸으로 간다.
            self.location_entry.focus_set()
        else:
            self.details.grid_remove()
            self.toggle_button.configure(text="상세")

    def _set_quick_fields(self, preview: WorkItem) -> None:
        """Keep the two most useful fields visible without duplicating state."""

        if preview.action == ACTION_WITH_DATE:
            self.quick_date_entry.configure(textvariable=self.at_var)
            self.quick_time_label.grid_remove()
            self.quick_time_entry.grid_remove()
        else:
            self.quick_date_entry.configure(textvariable=self.due_var)
            self.quick_time_label.grid()
            self.quick_time_entry.grid()

    # --- 파생 라벨 ------------------------------------------------------

    def _refresh(self, _event: tk.Event | None = None) -> None:
        if self.item.type == CALENDAR:
            has_start = bool(self.at_var.get().strip())
            if not has_start and not any(
                issue.code in {"missing_execution_time", "invalid_execution_time"}
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
                if issue.code in {"missing_execution_time", "invalid_execution_time"}:
                    issue.resolved = has_start
        has_due = bool(self.due_var.get().strip())
        for issue in self.item.review_issues:
            if issue.code in {"ambiguous_deadline_date", "invalid_deadline_date"}:
                issue.resolved = has_due
            if issue.code == "invalid_deadline_time":
                issue.resolved = bool(self.due_time_var.get().strip())
        if self.item.excluded_reason and self._current_signature() != self._excluded_signature:
            # 제목·일시·분류를 고치는 것은 중복을 해결하려는 명시적 수정이다.
            self.item.excluded_reason = None
            self._excluded_signature = None

        scope = self.scope_var.get()
        preview = WorkItem(
            type=CALENDAR if self.item.type == CALENDAR and not self.due_var.get().strip() else "todo",
            scope=scope,
            title=self.title() or "-",
            start=self.at_var.get().strip() or None,
            due=self.due_var.get().strip() or None,
            due_time=self.due_time_var.get().strip() or None,
        )
        label, hint = ACTION_BADGE[preview.action]
        self._set_quick_fields(preview)
        accent = theme.SCOPE_COLORS.get(scope, theme.MUTED)
        self.badge.configure(text=label, bg=accent, fg="#ffffff")
        self.badge_hint.configure(text=hint)
        self.strip.configure(bg=accent)
        summary = self._compact_summary(preview)
        self.compact_meta_var.set(summary)
        # 위 입력란이 이미 말한 것을 한 줄 더 쓰지 않는다.
        if summary:
            self.compact_meta_label.grid()
        else:
            self.compact_meta_label.grid_remove()

        if self.dismissed:
            self.selected_var.set(False)
            self._set_register_enabled(False)
            self.state_label.configure(text="미등록", foreground=theme.MUTED)
        elif self.item.excluded_reason:
            self.selected_var.set(False)
            self._set_register_enabled(False)
            self.state_label.configure(
                text=f"중복 제외 · {self.item.excluded_reason}", foreground=theme.WARN
            )
        elif scope == SCOPE_REFERENCE:
            self.selected_var.set(False)
            self._set_register_enabled(False)
            self.state_label.configure(text="참조 · 등록 안 함", foreground=theme.MUTED)
        elif self.item.saved:
            self.selected_var.set(False)
            self._set_register_enabled(False)
            self.state_label.configure(text="등록 완료", foreground=theme.OK)
        elif self.item.blocking_review_issues:
            self.selected_var.set(False)
            self._set_register_enabled(False)
            self.state_label.configure(text="확인 필요", foreground=theme.WARN)
        elif self.item.applicability == "unknown":
            self.selected_var.set(False)
            self._set_register_enabled(False)
            self.state_label.configure(text="대상 확인 필요", foreground=theme.WARN)
        elif self.item.source_state in {"completed", "informational"}:
            self.selected_var.set(False)
            self._set_register_enabled(False)
            self.state_label.configure(text="완료/안내 · 등록 안 함", foreground=theme.MUTED)
        elif self.item.applicability == "conditional":
            self._set_register_enabled(True)
            self.state_label.configure(text="조건부 · 선택 시 등록", foreground=theme.WARN)
        elif self.item.last_error:
            self._set_register_enabled(True)
            self.state_label.configure(text="등록 실패 · 다시 시도", foreground=theme.FAIL)
        else:
            self._set_register_enabled(True)
            self.state_label.configure(text="등록 대기", foreground=theme.MUTED)
        self._notify_change()

    def _compact_summary(self, preview: WorkItem) -> str:
        """입력란에 이미 보이는 날짜는 빼고, 남는 맥락만 한 줄로 모은다."""

        parts: list[str] = []
        at = self.at_var.get().strip()
        due = self.due_var.get().strip()
        if preview.action == ACTION_WITH_DATE and not at:
            parts.append("날짜 확인 필요")
        elif preview.action == ACTION_WITH_DUE and not due:
            parts.append("마감 확인 필요")
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
        return " · ".join(parts)

    def _current_signature(self) -> tuple[str, ...]:
        return (
            self.title(),
            self.at_var.get().strip(),
            self.due_var.get().strip(),
            self.scope_var.get().strip(),
        )

    # --- 값 꺼내기 ------------------------------------------------------

    def to_work_item(self) -> WorkItem:
        review_issues = list(self.item.review_issues)
        if self.at_var.get().strip():
            for issue in review_issues:
                if issue.code in {"missing_execution_time", "invalid_execution_time"}:
                    issue.resolved = True
        if self.due_var.get().strip():
            for issue in review_issues:
                if issue.code in {"ambiguous_deadline_date", "invalid_deadline_date"}:
                    issue.resolved = True
        if self.due_time_var.get().strip():
            for issue in review_issues:
                if issue.code == "invalid_deadline_time":
                    issue.resolved = True
        item = WorkItem(
            id=self.item.id,
            input_id=self.item.input_id,
            type="todo",
            scope=self.scope_var.get().strip(),
            title=self.title(),
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
            show_due_on_calendar=self.show_due_var.get(),
            calendar_entry_id=self.item.calendar_entry_id,
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

    def can_register_now(self) -> bool:
        """Evaluate the edited values, not only the original AI draft."""

        try:
            return self.to_work_item().can_register
        except ValueError:
            return False

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
        self._set_register_enabled(False)
        self.state_label.configure(text="등록 완료", foreground=theme.OK)

    def mark_excluded(self, reason: str) -> None:
        self.item.excluded_reason = reason
        self._excluded_signature = self._current_signature()
        self.selected_var.set(False)
        self._set_register_enabled(False)
        self.state_label.configure(text=f"중복 제외 · {reason}", foreground=theme.WARN)

    def mark_failed(self, error: str) -> None:
        # 실패한 카드는 다시 시도할 수 있어야 하므로 버튼 상태부터 되살린다.
        self.item.last_error = error
        self._refresh()
        self.state_label.configure(text=f"등록 실패 · {error}", foreground=theme.FAIL)
