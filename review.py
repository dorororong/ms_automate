"""Confirmation window: check the analysed items, edit them, save to Outlook."""

from __future__ import annotations

import json
from datetime import date
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable

import theme
from cards import WorkItemCard
from models import SCOPE_REFERENCE, WorkItem
from outlook_adapter import OutlookOperationError, OutlookUnavailableError
from service import Analysis, WorkflowService


class ReviewWindow(tk.Toplevel):
    """Display one completed analysis and collect a quick registration decision."""

    def __init__(
        self,
        master: tk.Misc,
        service: WorkflowService,
        *,
        on_closed: Callable[[], None] | None = None,
        on_finished: Callable[[str], None] | None = None,
        on_reanalyze: Callable[[str, bool, str | None], None] | None = None,
    ) -> None:
        super().__init__(master)
        self.service = service
        self.on_closed = on_closed
        self.on_finished = on_finished
        self.on_reanalyze = on_reanalyze
        self._finished = False
        self.analysis: Analysis | None = None
        self.cards: list[WorkItemCard] = []
        self.hidden_items: list[WorkItem] = []

        self.title("일정 확인 · Outlook")
        self.geometry("760x620")
        self.minsize(580, 440)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        theme.apply_theme(self)
        self.protocol("WM_DELETE_WINDOW", self.close)

        self.summary_var = tk.StringVar(value="분석 결과가 여기에 표시됩니다.")
        self.status_var = tk.StringVar(value="")
        self.source_var = tk.StringVar(value="")
        self.source_preview_var = tk.StringVar(value="")
        self.reference_var = tk.StringVar(value=date.today().isoformat())
        self._source_visible = True
        self._source_editable = False
        self._source_dirty = False
        self._context_controls_visible = False
        self.queue_var = tk.StringVar(value="")
        self._build_ui()
        self.bind("<Control-Return>", self._on_control_return)

    # --- construction -------------------------------------------------

    def _build_ui(self) -> None:
        P = theme.PAD_X
        head = theme.header(self, "등록 검토", "왼쪽 원문 · 오른쪽 저장 내용")
        head.grid(row=0, column=0, sticky="ew")
        ttk.Button(head, text="⋯", width=3, command=self._show_more_menu).grid(
            row=0, column=1, rowspan=2, padx=(8, 0), sticky="ne"
        )

        split = ttk.Frame(self, padding=(P, 0, P, 0))
        split.grid(row=1, column=0, sticky="nsew")
        split.columnconfigure(0, weight=1, uniform="review-pane")
        split.columnconfigure(1, weight=2, uniform="review-pane")
        split.rowconfigure(0, weight=1)

        # --- 왼쪽: 원문 -----------------------------------------------
        source = ttk.Frame(split, padding=(0, 0, 8, 0))
        source.grid(row=0, column=0, sticky="nsew")
        source.columnconfigure(0, weight=1)
        source.rowconfigure(1, weight=1)
        source_bar = ttk.Frame(source)
        source_bar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        source_bar.columnconfigure(0, weight=1)
        ttk.Label(source_bar, text="원문", style="Head.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        self.source_toggle_button = ttk.Button(
            source_bar, text="원문 수정", width=10, command=self._toggle_source
        )
        self.source_toggle_button.grid(row=0, column=1, sticky="e")

        text_wrap = tk.Frame(source, bg=theme.BORDER)
        text_wrap.grid(row=1, column=0, sticky="nsew")
        text_wrap.columnconfigure(0, weight=1)
        text_wrap.rowconfigure(0, weight=1)
        self.input_text = tk.Text(
            text_wrap, height=10, wrap="word", undo=True, font=theme.F_BODY,
            relief="flat", highlightthickness=0, padx=7, pady=6,
            background=theme.SURFACE, foreground=theme.TEXT,
            insertbackground=theme.TEXT, state="disabled",
        )
        self.input_text.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        input_scroll = ttk.Scrollbar(text_wrap, orient="vertical", command=self.input_text.yview)
        input_scroll.grid(row=0, column=1, sticky="ns", padx=(0, 1), pady=1)
        self.input_text.configure(yscrollcommand=input_scroll.set)
        self.input_text.bind("<KeyRelease>", self._on_source_changed)
        ttk.Label(
            source,
            textvariable=self.source_var,
            style="Muted.TLabel",
            anchor="w",
        ).grid(row=2, column=0, sticky="ew", pady=(6, 0))

        # 기준일과 분석 방식은 기본 화면을 가리지 않도록 메뉴/재분석에만 사용한다.
        self.source_detail = ttk.Frame(source)
        ttk.Label(self.source_detail, text="기준일", style="Muted.TLabel").pack(side="left")
        ttk.Entry(self.source_detail, textvariable=self.reference_var, width=11).pack(
            side="left", padx=(6, 8)
        )
        ttk.Label(self.source_detail, text="상대 날짜 기준", style="Muted.TLabel").pack(side="left")
        self.offline_button = ttk.Button(
            self.source_detail, text="오프라인", command=lambda: self.reanalyze(use_ai=False)
        )
        self.ai_button = ttk.Button(
            self.source_detail, text="다시 분석", command=lambda: self.reanalyze(use_ai=True)
        )

        # --- 오른쪽: 저장할 내용 -------------------------------------
        result = ttk.Frame(split, padding=(8, 0, 0, 0))
        result.grid(row=0, column=1, sticky="nsew")
        result.columnconfigure(0, weight=1)
        result.rowconfigure(1, weight=1)
        header_row = ttk.Frame(result)
        header_row.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        header_row.columnconfigure(0, weight=1)
        ttk.Label(header_row, text="저장할 내용", style="Head.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(header_row, textvariable=self.queue_var, style="Muted.TLabel").grid(
            row=0, column=1, sticky="e", padx=(6, 0)
        )
        ttk.Button(header_row, text="전체", width=5, command=lambda: self._set_all(True)).grid(
            row=0, column=2, padx=(6, 0)
        )
        ttk.Button(header_row, text="해제", width=5, command=lambda: self._set_all(False)).grid(
            row=0, column=3, padx=(4, 0)
        )

        self.copy_button = ttk.Button(self, text="JSON 복사", command=self._copy_json,
                                      state="disabled")
        self.export_button = ttk.Button(self, text="JSON 저장", command=self._export_json,
                                        state="disabled")

        canvas_frame = ttk.Frame(result)
        canvas_frame.grid(row=1, column=0, sticky="nsew")
        canvas_frame.columnconfigure(0, weight=1)
        canvas_frame.rowconfigure(0, weight=1)
        self.canvas = tk.Canvas(canvas_frame, highlightthickness=0, borderwidth=0,
                                background=theme.BG)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        canvas_scroll = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        canvas_scroll.grid(row=0, column=1, sticky="ns")
        self.canvas.configure(yscrollcommand=canvas_scroll.set)
        self.cards_frame = ttk.Frame(self.canvas)
        self.cards_frame.columnconfigure(0, weight=1)
        self._canvas_window = self.canvas.create_window((0, 0), window=self.cards_frame, anchor="nw")
        self.cards_frame.bind(
            "<Configure>",
            lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind(
            "<Configure>",
            lambda event: self.canvas.itemconfigure(self._canvas_window, width=event.width))
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)

        theme.divider(self).grid(row=2, column=0, sticky="ew", pady=(7, 0))
        bottom = ttk.Frame(self, padding=(P, 7, P, 8))
        bottom.grid(row=3, column=0, sticky="ew")
        bottom.columnconfigure(0, weight=1)
        ttk.Label(bottom, textvariable=self.status_var, style="Muted.TLabel").grid(
            row=0, column=0, sticky="w")
        ttk.Button(bottom, text="미등록", command=self.skip).grid(
            row=0, column=1, padx=(8, 0)
        )
        self.save_button = ttk.Button(
            bottom, text="등록", style="Accent.TButton",
            command=self.save_to_outlook, state="disabled")
        self.save_button.grid(row=0, column=2, padx=(8, 0))

    def _toggle_source(self) -> None:
        self._source_editable = not self._source_editable
        try:
            self.input_text.configure(state="normal" if self._source_editable else "disabled")
        except tk.TclError:
            return
        self.source_toggle_button.configure(
            text="수정 완료" if self._source_editable else "원문 수정"
        )

    def _on_source_changed(self, _event: tk.Event | None = None) -> None:
        if not self._source_editable:
            return
        self._source_dirty = True
        self.save_button.configure(state="disabled")
        self.status_var.set("원문이 바뀌었습니다. ⋯에서 다시 분석하세요.")

    def _set_source_visible(self, visible: bool) -> None:
        # Kept as a compatibility hook for callers from the earlier layout.
        # The source is intentionally always visible in the left pane.
        self._source_visible = True

    def _update_source_preview(self, text: str) -> None:
        compact = " ".join(text.split())
        if len(compact) > 72:
            compact = compact[:72].rstrip() + "…"
        self.source_preview_var.set(compact or "내용 없음")

    def _show_more_menu(self) -> None:
        """Keep optional analysis and export actions out of the main row."""

        menu = tk.Menu(
            self,
            tearoff=False,
            bg=theme.SURFACE,
            fg=theme.TEXT,
            activebackground=theme.ACCENT_SOFT,
            activeforeground=theme.TEXT,
            font=theme.F_BODY,
        )
        menu.add_command(
            label="원문 수정" if not self._source_editable else "원문 수정 완료",
            command=self._toggle_source,
        )
        menu.add_command(
            label="기준일 설정" if not self._context_controls_visible else "기준일 닫기",
            command=self._toggle_context_controls,
        )
        menu.add_command(label="다시 분석", command=lambda: self.request_reanalysis(use_ai=True))
        menu.add_command(label="오프라인 규칙으로 분석", command=lambda: self.request_reanalysis(use_ai=False))
        menu.add_separator()
        menu.add_command(
            label="JSON 복사",
            command=self._copy_json,
            state="normal" if self.analysis is not None else "disabled",
        )
        menu.add_command(
            label="JSON 저장",
            command=self._export_json,
            state="normal" if self.analysis is not None else "disabled",
        )
        try:
            menu.tk_popup(self.winfo_rootx() + self.winfo_width() - 170, self.winfo_rooty() + 34)
        finally:
            menu.grab_release()

    def _on_mousewheel(self, event: tk.Event) -> None:
        self.canvas.yview_scroll(int(-event.delta / 120), "units")

    def _toggle_context_controls(self) -> None:
        self._context_controls_visible = not self._context_controls_visible
        if self._context_controls_visible:
            self.source_detail.grid(row=3, column=0, sticky="ew", pady=(6, 0))
        else:
            self.source_detail.grid_remove()

    # --- lifecycle ----------------------------------------------------

    def present(self, text: str, *, use_ai: bool = True) -> None:
        """Compatibility entry point that delegates analysis to the app queue."""

        self._set_input_text(text)
        self.raise_window()
        if self.on_reanalyze is not None:
            self.request_reanalysis(use_ai=use_ai)
        else:
            self.status_var.set("앱의 분석 대기열에서 결과를 기다리는 중입니다.")

    def _set_input_text(self, text: str) -> None:
        self._source_editable = True
        self.input_text.configure(state="normal")
        self.input_text.delete("1.0", "end")
        self.input_text.insert("1.0", text)
        self.input_text.configure(state="disabled")
        self._source_editable = False
        self._source_dirty = False
        self.source_toggle_button.configure(text="원문 수정")
        self._update_source_preview(text)

    def present_analysis(self, analysis: Analysis) -> None:
        """Display a completed analysis; no LLM work occurs in this window."""

        self.analysis = analysis
        self._source_dirty = False
        self._set_input_text(analysis.text)
        self._show_analysis(analysis)
        self.raise_window()

    def raise_window(self) -> None:
        """Pull the window to the front; the hotkey fires from another app."""

        self.deiconify()
        self.lift()
        try:
            self.attributes("-topmost", True)
            self.after(400, self._drop_topmost)
            self.focus_force()
        except tk.TclError:
            pass

    def _drop_topmost(self) -> None:
        try:
            self.attributes("-topmost", False)
        except tk.TclError:
            pass

    def close(self) -> None:
        self._finish("skipped")

    def skip(self) -> None:
        """Leave the current result unregistered and show the next queued one."""

        self._finish("skipped")

    def _finish(self, decision: str) -> None:
        if self._finished:
            return
        self._finished = True
        callback = self.on_finished
        legacy_callback = self.on_closed
        try:
            if callback is not None:
                callback(decision)
            elif legacy_callback is not None:
                legacy_callback()
        finally:
            if self.winfo_exists():
                self.destroy()

    def _has_unsaved_items(self) -> bool:
        return any(
            not card.item.saved and not card.item.excluded_reason
            for card in self.cards
        )

    def _on_control_return(self, _event: tk.Event) -> str:
        self.save_to_outlook()
        return "break"

    # --- analysis -----------------------------------------------------

    def request_reanalysis(self, *, use_ai: bool) -> None:
        """Send explicit re-analysis back to the app queue.

        The review window never remains on screen while Solar Pro 4 is
        working. This is also how a manually edited source joins the same
        FIFO queue as new clipboard/file inputs.
        """

        text = self.input_text.get("1.0", "end-1c").strip()
        if not text:
            messagebox.showwarning("분석", "분석할 텍스트가 없습니다.", parent=self)
            return
        reference = self.reference_var.get().strip() or None
        if self.on_reanalyze is not None:
            self.on_reanalyze(text, use_ai, reference)
            self._finish("reanalysis")
            return
        self.status_var.set("분석 대기열에 연결된 창에서 다시 분석하세요.")

    def reanalyze(self, *, use_ai: bool) -> None:
        self.request_reanalysis(use_ai=use_ai)

    def _show_analysis(self, analysis: Analysis) -> None:
        self.status_var.set("기존 Outlook 내역과 겹치는지 확인합니다…")
        self.update_idletasks()
        try:
            conflicts = self.service.find_conflicts(analysis.work_items)
        except (
            AttributeError,
            OSError,
            ValueError,
            OutlookUnavailableError,
            OutlookOperationError,
        ) as exc:
            conflicts = {}
            warning = f"Outlook 중복 확인을 하지 못했습니다: {exc}"
            analysis.warning = f"{analysis.warning} · {warning}" if analysis.warning else warning
        for index, reason in conflicts.items():
            analysis.work_items[index].excluded_reason = reason
            analysis.work_items[index].selected = False

        self.analysis = analysis
        self.source_var.set(f"{analysis.model} · #{analysis.input_id}")
        self._set_source_visible(False)
        self._render_cards(analysis)

        if analysis.registerable:
            self.summary_var.set(analysis.summary)
            self.save_button.configure(state="normal")
            if analysis.ready_to_register:
                self.status_var.set("내용을 확인·수정한 뒤 등록할 항목을 체크하세요.")
            elif any(item.blocking_review_issues for item in analysis.registerable):
                self.status_var.set("확인이 필요한 항목은 수정 후 등록할 수 있습니다.")
            elif any(item.applicability == "unknown" for item in analysis.registerable):
                self.status_var.set("대상이 불명확한 항목은 확인 후 등록할 수 있습니다.")
            else:
                self.status_var.set("중복 항목은 제외했습니다. 제목·시간을 바꾸면 다시 등록할 수 있습니다.")
        elif analysis.work_items:
            self.summary_var.set(analysis.summary)
            self.save_button.configure(state="disabled")
            self.status_var.set("등록할 새 업무가 없습니다. 제외/참조 내용을 확인하세요.")
        else:
            self.summary_var.set("등록할 항목이 없습니다.")
            self.save_button.configure(state="disabled")
            self.status_var.set("참조로 분류된 내용만 있습니다.")

        self.copy_button.configure(state="normal")
        self.export_button.configure(state="normal")
        self._update_save_button()
        if analysis.warning:
            self.status_var.set(analysis.warning)

    def _set_all(self, checked: bool) -> None:
        for card in self.cards:
            if card.item.saved or card.item.excluded_reason or not card.item.can_register:
                card.selected_var.set(False)
            elif card.scope_var.get().strip() == SCOPE_REFERENCE:
                card.selected_var.set(False)
            else:
                card.selected_var.set(checked)
        self._update_save_button()

    def _render_cards(self, analysis: Analysis) -> None:
        for child in self.cards_frame.winfo_children():
            child.destroy()
        self.cards = []
        self.hidden_items = []

        visible_items: list[WorkItem] = []
        for item in analysis.work_items:
            if (
                item.scope == SCOPE_REFERENCE
                or item.source_state in {"completed", "informational"}
                or item.excluded_reason
            ):
                self.hidden_items.append(item)
            else:
                visible_items.append(item)

        if not visible_items:
            ttk.Label(self.cards_frame, text="분석 결과가 없습니다.", style="Muted.TLabel").grid(
                row=0, column=0, padx=4, pady=8, sticky="w"
            )
        for index, item in enumerate(visible_items, start=1):
            card = WorkItemCard(
                self.cards_frame,
                item,
                index,
                on_change=self._update_save_button,
            )
            card.frame.grid(row=index, column=0, sticky="ew", padx=(0, 4), pady=(0, 8))
            self.cards.append(card)

        ignored_lines = list(analysis.ignored)
        for item in self.hidden_items:
            if item.title in ignored_lines:
                continue
            if item.source_state in {"completed", "informational"}:
                label = "완료/안내"
            elif item.excluded_reason:
                label = f"중복 제외 · {item.excluded_reason}"
            else:
                label = "참조"
            ignored_lines.append(f"{item.title} ({label})")

        if ignored_lines:
            ignored_row = ttk.Frame(self.cards_frame)
            ignored_row.grid(
                row=len(self.cards) + 1,
                column=0,
                sticky="ew",
                padx=4,
                pady=(2, 10),
            )
            ignored_row.columnconfigure(0, weight=1)
            ignored_details = ttk.Label(
                ignored_row,
                text="\n".join(f"- {line}" for line in ignored_lines),
                style="Muted.TLabel",
                justify="left",
            )
            ignored_details.grid(row=1, column=0, sticky="w", pady=(4, 0))
            ignored_details.grid_remove()

            def toggle_ignored() -> None:
                if ignored_details.winfo_manager():
                    ignored_details.grid_remove()
                    ignored_toggle.configure(text=f"참조/제외 {len(ignored_lines)}건 보기")
                else:
                    ignored_details.grid()
                    ignored_toggle.configure(text="참조/제외 접기")

            ignored_toggle = ttk.Button(
                ignored_row,
                text=f"참조/제외 {len(ignored_lines)}건 보기",
                command=toggle_ignored,
            )
            ignored_toggle.grid(row=0, column=0, sticky="w")

        self.canvas.yview_moveto(0.0)
        self._update_save_button()

    def _update_save_button(self) -> None:
        if not hasattr(self, "save_button"):
            return
        count = sum(
            1
            for card in self.cards
            if card.selected_var.get()
            and card.item.can_register
            and not card.item.saved
            and not card.item.excluded_reason
        )
        self.save_button.configure(text=f"{count}건 등록" if count else "등록")

    def set_queue_status(self, count: int) -> None:
        """Show pending analyses without adding another panel to the UI."""

        self.queue_var.set(f"대기 {count}건" if count else "")

    # --- recording ----------------------------------------------------

    def save_to_outlook(self) -> None:
        if self._source_dirty:
            messagebox.showwarning(
                "Outlook 등록",
                "원문이 변경되었습니다. ⋯에서 다시 분석한 뒤 등록하세요.",
                parent=self,
            )
            return
        if self.analysis is None or not self.cards:
            messagebox.showwarning("Outlook 등록", "먼저 텍스트를 분석하세요.", parent=self)
            return
        try:
            drafts = [card.to_work_item() for card in self.cards]
        except ValueError as exc:
            messagebox.showerror("Outlook 등록", str(exc), parent=self)
            return
        if not any(item.selected and not item.saved for item in drafts):
            messagebox.showinfo("Outlook 등록", "등록할 항목을 선택하세요.", parent=self)
            return

        self.save_button.configure(state="disabled")
        self.status_var.set("현재 로그인된 Outlook에 등록하는 중...")
        self.update_idletasks()

        # Outlook COM is apartment-threaded; this stays on the Tk thread.
        outcome = self.service.record(drafts)

        for index, (card, item) in enumerate(zip(self.cards, drafts)):
            if index in outcome.duplicates:
                card.mark_excluded(outcome.duplicates[index])
            elif index in outcome.errors:
                card.mark_failed(outcome.errors[index])
            elif item.saved and not card.item.saved:
                card.mark_saved(item.outlook_entry_id)

        self.save_button.configure(state="normal")
        parts = [f"Outlook 등록 {outcome.saved}개"]
        if outcome.skipped:
            parts.append(f"참조 {outcome.skipped}개는 로컬 기록만")
        if outcome.failed:
            parts.append(f"실패 {outcome.failed}개 — 수정 후 다시 시도하세요")
        if outcome.duplicates:
            parts.append(f"중복 제외 {len(outcome.duplicates)}개")
        if outcome.blocked:
            parts.append("중복 확인 실패로 등록을 보류했습니다")
        if outcome.warning:
            parts.append(outcome.warning)
        self.status_var.set(" · ".join(parts))
        self._update_save_button()
        if outcome.failed == 0 and not outcome.blocked:
            self._finish("registered")

    # --- JSON ---------------------------------------------------------

    def _payload(self) -> dict[str, Any]:
        analysis = self.analysis
        if analysis is None:
            return {"schema_version": 1, "items": [], "ignored": []}
        card_by_item = {id(card.item): card for card in self.cards}
        drafts: list[WorkItem] = []
        for item in analysis.work_items:
            card = card_by_item.get(id(item))
            if card is None:
                drafts.append(item)
                continue
            try:
                drafts.append(card.to_work_item())
            except ValueError:
                drafts.append(item)
        return {
            "schema_version": 2,
            "input_id": analysis.input_id,
            "source": analysis.source,
            "model": analysis.model,
            "context": analysis.context,
            "items": [
                {
                    "scope": item.scope,
                    "action": item.action,
                    "type": item.type,
                    "title": item.title,
                    "start": item.start,
                    "end": item.end,
                    "due": item.due,
                    "due_time": item.due_time,
                    "description": item.description,
                    "location": item.location,
                    "category": item.category,
                    "all_day": item.all_day,
                    "reminder_minutes": item.reminder_minutes,
                    "reminder": item.reminder,
                    "repeat_freq": item.repeat_freq,
                    "repeat_detail": item.repeat_detail,
                    "saved": item.saved,
                    "outlook_entry_id": item.outlook_entry_id,
                    "excluded_reason": item.excluded_reason,
                    "intent": item.intent,
                    "applicability": item.applicability,
                    "condition": item.condition,
                    "source_state": item.source_state,
                    "temporal_context": [
                        context.to_dict() for context in item.temporal_context
                    ],
                    "checklist": list(item.checklist),
                    "evidence": [evidence.to_dict() for evidence in item.evidence],
                    "review_issues": [issue.to_dict() for issue in item.review_issues],
                    "group_id": item.group_id,
                }
                for item in drafts
            ],
            "ignored": analysis.ignored,
        }

    def _copy_json(self) -> None:
        output = json.dumps(self._payload(), ensure_ascii=False, indent=2)
        self.clipboard_clear()
        self.clipboard_append(output)
        self.status_var.set("JSON을 클립보드에 복사했습니다.")

    def _export_json(self) -> None:
        path = filedialog.asksaveasfilename(
            parent=self,
            title="분석 결과 JSON 저장",
            defaultextension=".json",
            filetypes=[("JSON", "*.json"), ("모든 파일", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as handle:
                json.dump(self._payload(), handle, ensure_ascii=False, indent=2)
        except OSError as exc:
            messagebox.showerror("JSON 저장 실패", str(exc), parent=self)
            return
        self.status_var.set(f"JSON을 저장했습니다: {path}")
