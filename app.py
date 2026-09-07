"""Compact CATMOA-style launcher for the text/file → Solar Pro 4 → Outlook flow.

The main window is intentionally focused on one short path:

    메시지 붙여넣기/선택 → 고양이에게 파일 드롭 → AI 추출 → 중복 제외 → 빠른 등록

Outlook item management is reserved for a separate future UI and is not
exposed in this registration flow.
"""

from __future__ import annotations

from collections import deque
from dataclasses import dataclass
from tkinter import filedialog, messagebox
import queue
import threading
import tkinter as tk
from tkinter import ttk
from pathlib import Path

import selection
import theme
from file_input import FileLoadError, FileLoadResult, load_files
from hotkey import HotkeyError, HotkeyListener
from outlook_adapter import OutlookUnavailableError
from review import ReviewWindow
from service import Analysis, WorkflowService
from windows_drop import WindowsFileDropHandler

try:
    from tkinterdnd2 import DND_FILES, DND_TEXT, TkinterDnD
    _TK_ROOT = TkinterDnD.Tk
    _TK_DND_AVAILABLE = True
except ImportError:  # Native WM_DROPFILES still covers file drops.
    DND_FILES = DND_TEXT = None
    _TK_ROOT = tk.Tk
    _TK_DND_AVAILABLE = False


HOTKEY_LABEL = "Ctrl+Shift+X"
VK_X = 0x58
TASKBAR_MARGIN = 72
WINDOW_WIDTH = 350
WINDOW_HEIGHT = 320

FACE = {
    "idle": "(=^ω^=)",
    "eating": "(=^o^=)",
    "done": "(=^▽^=)",
    "error": "(=；ω；=)",
}


@dataclass(frozen=True)
class AnalysisRequest:
    """One FIFO unit. File reading happens before this request is created."""

    text: str
    use_ai: bool = True
    reference_date: str | None = None
    source: str = "text"


class MiniWindow(_TK_ROOT):
    """Friendly always-on-top drop panel that starts the registration flow."""

    def __init__(self) -> None:
        super().__init__()
        self.service = WorkflowService()
        self.review: ReviewWindow | None = None
        self._drop_handler: WindowsFileDropHandler | None = None
        self._dnd_enabled = False

        self._events: queue.Queue[tuple[str, object]] = queue.Queue()
        self._analysis_jobs: queue.Queue[AnalysisRequest | None] = queue.Queue()
        self._ready_analyses: deque[Analysis] = deque()
        self._analysis_errors: deque[str] = deque()
        self._analysis_in_flight = 0
        self._analysis_thread = threading.Thread(
            target=self._analysis_worker,
            name="analysis-queue",
            daemon=True,
        )
        self._analysis_thread.start()
        self._capture_lock = threading.Lock()
        self.listener = HotkeyListener(self._on_hotkey, key_code=VK_X, label=HOTKEY_LABEL)

        self.title("업무 정리 · Outlook")
        self.geometry(self._corner_geometry())
        self.resizable(False, False)
        self.attributes("-topmost", True)
        self.protocol("WM_DELETE_WINDOW", self.quit_app)
        theme.apply_theme(self)

        self.status_var = tk.StringVar(value="메시지를 넣어 주세요.")
        self._build_ui()
        self.after_idle(self._install_file_drop)
        self._start_hotkey()
        self.after(100, self._drain_events)

    # --- layout -------------------------------------------------------

    def _corner_geometry(self) -> str:
        x = max(self.winfo_screenwidth() - WINDOW_WIDTH - 16, 0)
        y = max(self.winfo_screenheight() - WINDOW_HEIGHT - TASKBAR_MARGIN, 0)
        return f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{x}+{y}"

    def _build_ui(self) -> None:
        P = theme.PAD_X
        head = theme.header(self, "업무 정리", "붙여넣거나 파일을 놓으세요")
        head.pack(fill="x")
        self.more_button = ttk.Button(
            head, text="⋯", width=3, command=self._show_more_menu,
        )
        self.more_button.grid(row=0, column=1, rowspan=2, padx=(8, 0), sticky="ne")

        # 파일 입력구는 작은 고정 영역으로 두고, 세부 형식 안내는 화면에서 뺀다.
        self.drop_zone = tk.Frame(
            self,
            bg=theme.ACCENT_SOFT,
            highlightbackground=theme.BORDER,
            highlightthickness=1,
            cursor="hand2",
        )
        self.drop_zone.pack(fill="x", padx=P, pady=(0, 8))
        self.drop_zone.configure(height=70)
        self.drop_zone.pack_propagate(False)
        self.drop_zone.columnconfigure(0, weight=1)
        self.face_var = tk.StringVar(value=FACE["idle"])
        self.face_label = tk.Label(
            self.drop_zone,
            textvariable=self.face_var,
            font=("Consolas", 18, "bold"),
            bg=theme.ACCENT_SOFT,
            fg=theme.ACCENT,
        )
        self.face_label.grid(row=0, column=0, pady=(5, 0))
        self.drop_title_var = tk.StringVar(value="파일을 놓거나 클릭")
        self.drop_title = tk.Label(
            self.drop_zone,
            textvariable=self.drop_title_var,
            font=theme.F_HEAD,
            bg=theme.ACCENT_SOFT,
            fg=theme.TEXT,
        )
        self.drop_title.grid(row=1, column=0)
        self.drop_hint = tk.Label(
            self.drop_zone,
            text="파일을 놓거나 클릭해서 선택",
            font=theme.F_SMALL,
            bg=theme.ACCENT_SOFT,
            fg=theme.MUTED,
            justify="center",
        )
        self.drop_hint.grid(row=2, column=0, pady=(1, 0))
        for widget in (self.drop_zone, self.face_label, self.drop_title, self.drop_hint):
            widget.bind("<Button-1>", lambda _event: self.choose_files())
            widget.bind("<Enter>", self._drop_hover)
            widget.bind("<Leave>", self._drop_leave)

        # Tk Text는 일반 복사·붙여넣기와 선택 드래그를 그대로 지원한다.
        text_section = ttk.Frame(self, padding=(P, 0, P, 0))
        text_section.pack(fill="x")
        text_section.columnconfigure(0, weight=1)
        input_head = ttk.Frame(text_section)
        input_head.grid(row=0, column=0, sticky="ew", pady=(0, 3))
        ttk.Label(input_head, text="메시지", style="Head.TLabel").pack(side="left")
        ttk.Label(input_head, text="Ctrl+V", style="Muted.TLabel").pack(side="right")
        text_wrap = tk.Frame(text_section, bg=theme.BORDER)
        text_wrap.grid(row=1, column=0, sticky="ew")
        text_wrap.columnconfigure(0, weight=1)
        self.input_text = tk.Text(
            text_wrap,
            width=34,
            height=3,
            wrap="word",
            undo=True,
            font=theme.F_BODY,
            relief="flat",
            highlightthickness=0,
            padx=7,
            pady=5,
            background=theme.SURFACE,
            foreground=theme.TEXT,
        )
        self.input_text.grid(row=0, column=0, sticky="ew", padx=1, pady=1)
        input_scroll = ttk.Scrollbar(text_wrap, orient="vertical", command=self.input_text.yview)
        input_scroll.grid(row=0, column=1, sticky="ns", padx=(0, 1), pady=1)
        self.input_text.configure(yscrollcommand=input_scroll.set)
        self.input_text.bind("<Control-Return>", self._analyze_input)

        actions = ttk.Frame(text_section)
        actions.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        actions.columnconfigure(0, weight=1)
        actions.columnconfigure(1, weight=1)
        ttk.Button(
            actions,
            text="일정 찾기",
            style="Accent.TButton",
            command=self._analyze_input,
        ).grid(row=0, column=0, sticky="ew", padx=(0, 3))
        ttk.Button(
            actions,
            text="파일",
            command=self.choose_files,
        ).grid(row=0, column=1, sticky="ew", padx=(3, 0))

        bottom = ttk.Frame(self, padding=(P, 7, P, 8))
        bottom.pack(fill="x")
        bottom.columnconfigure(0, weight=1)
        ttk.Label(
            bottom,
            textvariable=self.status_var,
            style="Status.TLabel",
            anchor="w",
            wraplength=WINDOW_WIDTH - 110,
        ).grid(row=0, column=0, sticky="ew")

    # --- drop/input ---------------------------------------------------

    def _analysis_worker(self) -> None:
        """Run Solar/Offline analysis sequentially in FIFO order."""

        while True:
            request = self._analysis_jobs.get()
            if request is None:
                return
            try:
                analysis = self.service.prepare(
                    request.text,
                    use_ai=request.use_ai,
                    reference_date=request.reference_date,
                    reference_source="user_or_today",
                    timezone_name="Asia/Seoul",
                    input_source=request.source,
                )
                self._events.put(("analysis_ready", analysis))
            except Exception as exc:
                self._events.put(("analysis_error", str(exc)))

    def _enqueue_analysis(
        self,
        text: str,
        *,
        use_ai: bool = True,
        reference_date: str | None = None,
        source: str = "text",
    ) -> None:
        clean = text.strip()
        if not clean:
            return
        self._analysis_in_flight += 1
        self._analysis_jobs.put(
            AnalysisRequest(
                text=clean,
                use_ai=use_ai,
                reference_date=reference_date,
                source=source,
            )
        )
        self._set_cat("eating", "분석 대기…")
        self._update_queue_status()

    def _queue_count(self) -> int:
        return self._analysis_in_flight + len(self._ready_analyses)

    def _update_queue_status(self) -> None:
        pending = self._queue_count()
        if self.review is not None and self.review.winfo_exists():
            self.review.set_queue_status(pending)
            if pending:
                self.status_var.set(f"현재 항목 확인 중 · 다음 {pending}건 대기")
            return
        if pending:
            self.status_var.set(
                "분석 중…" if self._analysis_in_flight else f"검토 대기 {pending}건"
            )

    def _set_cat(self, state: str, title: str | None = None) -> None:
        self.face_var.set(FACE.get(state, FACE["idle"]))
        if title:
            self.drop_title_var.set(title)
        elif state == "idle":
            self.drop_title_var.set("파일을 놓거나 클릭")
        elif state == "eating":
            self.drop_title_var.set("읽는 중…")
        elif state == "done":
            self.drop_title_var.set("읽었어요")
        elif state == "error":
            self.drop_title_var.set("읽지 못했어요")

    def _drop_hover(self, _event: tk.Event) -> None:
        hover = "#dbeefa"
        self.drop_zone.configure(bg=hover)
        for widget in (self.face_label, self.drop_title, self.drop_hint):
            widget.configure(bg=hover)

    def _drop_leave(self, _event: tk.Event) -> None:
        self.drop_zone.configure(bg=theme.ACCENT_SOFT)
        for widget in (self.face_label, self.drop_title, self.drop_hint):
            widget.configure(bg=theme.ACCENT_SOFT)

    def _install_file_drop(self) -> None:
        if _TK_DND_AVAILABLE:
            try:
                self.drop_target_register(DND_FILES, DND_TEXT)
                self.dnd_bind("<<Drop>>", self._on_tk_drop)
                self._dnd_enabled = True
                return
            except (AttributeError, RuntimeError, tk.TclError):
                self._dnd_enabled = False
        try:
            self._drop_handler = WindowsFileDropHandler(self, self._on_files_dropped)
        except (OSError, RuntimeError, TypeError) as exc:
            # The file picker remains available if native window subclassing
            # is blocked by a particular Windows/Tk build.
            self.status_var.set(f"파일 드롭을 준비하지 못했습니다. 파일 선택을 사용하세요. ({exc})")

    def _on_files_dropped(self, paths: list[str]) -> None:
        self._load_files_async(paths)

    def _on_tk_drop(self, event: object) -> None:
        """Handle tkinterdnd2's file and plain-text drop formats."""

        raw = str(getattr(event, "data", "")).strip()
        if not raw:
            return
        try:
            tokens = list(self.tk.splitlist(raw))
        except tk.TclError:
            tokens = []
        if tokens and all(Path(token).is_file() for token in tokens):
            self._on_files_dropped(tokens)
            return
        self.input_text.delete("1.0", "end")
        self.input_text.insert("1.0", raw)
        self._analyze_input()

    def choose_files(self) -> None:
        paths = filedialog.askopenfilenames(
            parent=self,
            title="일정이 들어 있는 파일 선택",
            filetypes=[
                ("일정 문서", "*.pdf *.hwp *.hwpx *.docx *.txt *.md *.csv *.tsv"),
                ("이미지", "*.png *.jpg *.jpeg *.bmp *.gif *.webp *.tiff"),
                ("모든 파일", "*.*"),
            ],
        )
        if paths:
            self._load_files_async(list(paths))

    def _load_files_async(self, paths: list[str]) -> None:
        self._set_cat("eating")
        self.status_var.set(f"{len(paths)}개 파일 읽는 중…")

        def worker() -> None:
            try:
                result = load_files(paths)
            except FileLoadError as exc:
                self._events.put(("files_error", str(exc)))
                return
            self._events.put(("files_loaded", result))

        threading.Thread(target=worker, name="file-loader", daemon=True).start()

    def _analyze_input(self, _event: tk.Event | None = None) -> str | None:
        text = self.input_text.get("1.0", "end-1c").strip()
        if not text:
            messagebox.showwarning("일정 찾기", "메시지를 넣어 주세요.", parent=self)
            self.input_text.focus_set()
            return "break" if _event is not None else None
        self._enqueue_analysis(text, source="text")
        return "break" if _event is not None else None

    # --- hotkey -------------------------------------------------------

    def _start_hotkey(self) -> None:
        try:
            self.listener.start()
        except HotkeyError as exc:
            self.status_var.set(f"{exc} 파일 선택·붙여넣기로 계속 사용할 수 있습니다.")

    def _on_hotkey(self) -> None:
        """Runs on the hotkey thread. Capture here, hand text to the UI."""

        if not self._capture_lock.acquire(blocking=False):
            return
        try:
            picked = selection.capture_selection(release_keys=(VK_X,))
        except Exception as exc:
            self._events.put(("error", f"텍스트를 가져오지 못했습니다: {exc}"))
            return
        finally:
            self._capture_lock.release()

        if not picked.text:
            self._events.put((
                "error",
                "선택한 텍스트도 클립보드 내용도 없습니다. 텍스트를 선택한 뒤 다시 누르세요.",
            ))
            return
        self._events.put((picked.source, picked.text))

    def _drain_events(self) -> None:
        try:
            while True:
                kind, payload = self._events.get_nowait()
                if kind == "error":
                    self._set_cat("error")
                    self.status_var.set(str(payload))
                    self.bell()
                elif kind == "files_error":
                    self._set_cat("error")
                    self.status_var.set("파일을 읽지 못했습니다.")
                    messagebox.showerror("파일 읽기 실패", str(payload), parent=self)
                elif kind == "files_loaded":
                    self._on_files_loaded(payload)
                elif kind == "analysis_ready":
                    self._on_analysis_ready(payload)
                elif kind == "analysis_error":
                    self._on_analysis_error(str(payload))
                else:
                    text = str(payload)
                    self.input_text.delete("1.0", "end")
                    self.input_text.insert("1.0", text)
                    self._enqueue_analysis(
                        text,
                        source="selection" if kind == "selection" else "clipboard",
                    )
        except queue.Empty:
            pass
        finally:
            if self.winfo_exists():
                self.after(100, self._drain_events)

    def _on_files_loaded(self, result: object) -> None:
        if not isinstance(result, FileLoadResult):
            self._set_cat("error")
            self.status_var.set("파일 처리 결과를 읽지 못했습니다.")
            return
        self.input_text.delete("1.0", "end")
        self.input_text.insert("1.0", result.text)
        self._set_cat("done")
        if result.warnings:
            self.status_var.set("일부 파일을 읽지 못했어요.")
        else:
            self.status_var.set(f"{len(result.names)}개 파일을 읽었어요.")
        self._enqueue_analysis(result.text, source="file")

    def capture_now(self) -> None:
        """Same path as the hotkey, for when the shortcut is unavailable."""

        threading.Thread(target=self._on_hotkey, name="manual-capture", daemon=True).start()

    # --- windows ------------------------------------------------------

    def _review_window(self) -> ReviewWindow:
        if self.review is None or not self.review.winfo_exists():
            self.review = ReviewWindow(
                self,
                self.service,
                on_finished=self._on_review_finished,
                on_reanalyze=self._on_reanalyze,
            )
        return self.review

    def _on_reanalyze(
        self, text: str, use_ai: bool, reference_date: str | None
    ) -> None:
        self._enqueue_analysis(
            text,
            use_ai=use_ai,
            reference_date=reference_date,
            source="reanalysis",
        )

    def _on_review_finished(self, _decision: str) -> None:
        self.review = None
        self.after_idle(self._show_next_analysis)

    def _open_review(self, text: str) -> None:
        self._enqueue_analysis(text, source="text")

    def _on_analysis_ready(self, payload: object) -> None:
        if self._analysis_in_flight > 0:
            self._analysis_in_flight -= 1
        if isinstance(payload, Analysis):
            self._ready_analyses.append(payload)
        else:
            self._analysis_errors.append("분석 결과를 읽지 못했습니다.")
        self._show_next_analysis()

    def _on_analysis_error(self, error: str) -> None:
        if self._analysis_in_flight > 0:
            self._analysis_in_flight -= 1
        self._analysis_errors.append(error)
        if self.review is None or not self.review.winfo_exists():
            self._set_cat("error")
            self.status_var.set("분석 실패 · 다음 항목을 준비합니다.")
            self.bell()
        self._show_next_analysis()

    def _show_next_analysis(self) -> None:
        if self.review is not None and self.review.winfo_exists():
            self._update_queue_status()
            return
        if self._ready_analyses:
            analysis = self._ready_analyses.popleft()
            self.attributes("-topmost", False)
            review = self._review_window()
            review.present_analysis(analysis)
            self._update_queue_status()
            return
        if self._analysis_in_flight:
            self._set_cat("eating", "분석 중…")
            self.status_var.set("분석 중…")
            return
        if self._analysis_errors:
            error = self._analysis_errors.popleft()
            self._set_cat("error")
            self.status_var.set(f"분석 실패: {error[:80]}")
            self.bell()
            return
        self._set_cat("idle")
        self.status_var.set("메시지를 넣어 주세요.")
        self.attributes("-topmost", True)

    def open_blank_review(self) -> None:
        self.input_text.focus_set()
        self.status_var.set("메시지를 넣어 주세요.")

    def _show_more_menu(self) -> None:
        """Keep secondary actions available without crowding the launcher."""

        menu = tk.Menu(
            self,
            tearoff=False,
            bg=theme.SURFACE,
            fg=theme.TEXT,
            activebackground=theme.ACCENT_SOFT,
            activeforeground=theme.TEXT,
            font=theme.F_BODY,
        )
        mode = "Solar Pro 4 AI" if self.service.has_api_key() else "오프라인 규칙"
        menu.add_command(label=f"분석 방식 · {mode}", state="disabled")
        menu.add_separator()
        menu.add_command(label=f"{HOTKEY_LABEL} 선택 텍스트 가져오기", command=self.capture_now)
        menu.add_command(label="Outlook 연결 확인", command=self.test_outlook)
        menu.add_command(label="단축키 안내", command=self._show_shortcuts)
        try:
            menu.tk_popup(self.winfo_rootx() + self.winfo_width() - 150, self.winfo_rooty() + 32)
        finally:
            menu.grab_release()

    def _show_shortcuts(self) -> None:
        messagebox.showinfo(
            "사용 방법",
            f"{HOTKEY_LABEL}: 선택한 텍스트 가져오기\n"
            "Ctrl+V: 메시지 붙여넣기\n"
            "Ctrl+Enter: 일정 찾기",
            parent=self,
        )

    def test_outlook(self) -> None:
        try:
            user_name = self.service.test_outlook()
        except OutlookUnavailableError as exc:
            self.status_var.set("클래식 Outlook 연결 실패")
            messagebox.showerror("Outlook 연결 실패", str(exc), parent=self)
            return
        self.status_var.set(f"Outlook 연결됨: {user_name}")

    # --- shutdown -----------------------------------------------------

    def quit_app(self) -> None:
        if self._drop_handler is not None:
            self._drop_handler.close()
        self.listener.stop()
        self._analysis_jobs.put(None)
        self.destroy()


def main() -> None:
    MiniWindow().mainloop()


if __name__ == "__main__":
    main()
