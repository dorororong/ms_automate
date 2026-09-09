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

import icons
import selection
import theme
from file_input import FileLoadError, FileLoadResult, load_files
from hotkey import HotkeyError, HotkeyListener
from outlook_adapter import OutlookUnavailableError
from review import ReviewWindow
from service import Analysis, WorkflowService
from tray import TrayError, TrayIcon
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
WINDOW_WIDTH = 120
WINDOW_HEIGHT = 160

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
        self._paused = False
        self.tray: TrayIcon | None = None

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
        self.protocol("WM_DELETE_WINDOW", self._on_close_button)
        # default=True 라서 검토창을 포함한 모든 창에 같은 아이콘이 붙는다.
        icons.apply_window_icon(self)
        theme.apply_theme(self)

        self.status_var = tk.StringVar(value="대기 중")
        self.detail_var = tk.StringVar(value=f"파일을 놓아 주세요 · {HOTKEY_LABEL}")
        self.mirror_var = tk.StringVar(value="동기화 준비 중")
        self.mirror_detail = "Outlook 동기화를 준비하고 있습니다."
        self._build_ui()
        # 입력창을 없앤 뒤에도 붙여넣기 경로는 남긴다.
        self.bind("<Control-v>", self.paste_from_clipboard)
        self.bind("<Control-V>", self.paste_from_clipboard)
        self.after_idle(self._install_file_drop)
        self._start_hotkey()
        self._start_tray()
        # 중복 검사가 COM 대신 읽을 로컬 미러를 백그라운드에서 채운다.
        self.service.mirror.start(self._on_mirror_event)
        self.after(100, self._drain_events)

    # --- layout -------------------------------------------------------

    def _corner_geometry(self) -> str:
        x = max(self.winfo_screenwidth() - WINDOW_WIDTH - 16, 0)
        y = max(self.winfo_screenheight() - WINDOW_HEIGHT - TASKBAR_MARGIN, 0)
        return f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{x}+{y}"

    def _build_ui(self) -> None:
        """상태 하나만 보여주는 창.

        입력은 전역 단축키·파일 드롭·`⋯` 메뉴로 들어오므로, 창에는 지금 무슨 일이
        일어나고 있는지만 남깁니다. 버튼과 입력창을 걷어낸 자리에 상태를 크게 둡니다.
        """

        head = ttk.Frame(self, padding=(8, 6, 6, 4))
        head.pack(fill="x")
        head.columnconfigure(0, weight=1)
        ttk.Label(head, text="업무 정리", style="Head.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        self.more_button = ttk.Button(
            head, text="⋯", width=2, command=self._show_more_menu,
        )
        self.more_button.grid(row=0, column=1, sticky="e")

        # 창 전체가 드롭 구역이자 상태 표시판이다.
        self.drop_zone = tk.Frame(
            self,
            bg=theme.ACCENT_SOFT,
            highlightbackground=theme.BORDER,
            highlightthickness=1,
            cursor="hand2",
        )
        self.drop_zone.pack(fill="both", expand=True, padx=8, pady=(0, 4))
        self.drop_zone.pack_propagate(False)
        self.drop_zone.columnconfigure(0, weight=1)
        self.drop_zone.rowconfigure(0, weight=1)
        self.drop_zone.rowconfigure(3, weight=1)
        self.face_var = tk.StringVar(value=FACE["idle"])
        self.face_label = tk.Label(
            self.drop_zone,
            textvariable=self.face_var,
            font=("Consolas", 12, "bold"),
            bg=theme.ACCENT_SOFT,
            fg=theme.ACCENT,
        )
        self.face_label.grid(row=1, column=0)
        self.state_label = tk.Label(
            self.drop_zone,
            textvariable=self.status_var,
            font=theme.F_HEAD,
            bg=theme.ACCENT_SOFT,
            fg=theme.TEXT,
            wraplength=WINDOW_WIDTH - 26,
            justify="center",
        )
        self.state_label.grid(row=2, column=0, pady=(2, 0))
        for widget in (self.drop_zone, self.face_label, self.state_label):
            widget.bind("<Button-1>", lambda _event: self.choose_files())
            widget.bind("<Enter>", self._drop_hover)
            widget.bind("<Leave>", self._drop_leave)

        bottom = ttk.Frame(self, padding=(8, 0, 8, 5))
        bottom.pack(fill="x")
        bottom.columnconfigure(0, weight=1)
        ttk.Label(
            bottom,
            textvariable=self.mirror_var,
            style="Muted.TLabel",
            anchor="w",
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
        if self._paused:
            self._set_cat("idle")
            self.status_var.set("일시중지 중이라 넣지 않았습니다.")
            self.bell()
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
        self._update_queue_status()

    def _queue_count(self) -> int:
        return self._analysis_in_flight + len(self._ready_analyses)

    # --- 트레이 -----------------------------------------------------------

    def _start_tray(self) -> None:
        """작은 창을 닫아도 앱을 멈추거나 끌 수 있는 자리를 만든다."""

        tray = TrayIcon(
            "업무 정리 · Outlook",
            on_show=lambda: self._events.put(("tray_show", None)),
            on_toggle_pause=lambda: self._events.put(("tray_pause", None)),
            on_quit=lambda: self._events.put(("tray_quit", None)),
            pause_label=lambda: "재개" if self._paused else "일시중지",
        )
        try:
            tray.start()
        except TrayError as exc:
            self.mirror_detail = f"트레이 아이콘 없음: {str(exc)[:60]}"
            return
        self.tray = tray

    def _sync_tray_tooltip(self) -> None:
        # 툴팁은 이름만 짧게 둔다. 상태는 창과 `⋯` 메뉴에 있다.
        if self.tray is None:
            return
        self.tray.set_tooltip("업무 정리")

    def _on_close_button(self) -> None:
        """X 는 트레이로 내리고, 종료는 트레이나 `⋯` 메뉴에서 한다.

        트레이 아이콘을 만들지 못한 환경에서는 창이 유일한 조작 수단이므로
        예전처럼 그대로 종료합니다.
        """

        if self.tray is None:
            self.quit_app()
            return
        self.withdraw()

    def show_window(self) -> None:
        """트레이에서 부를 때 창을 다시 앞으로 가져온다."""

        try:
            self.deiconify()
            self.lift()
            self.focus_force()
        except tk.TclError:
            pass

    def toggle_pause(self) -> None:
        """멈춘 동안에는 단축키·드롭으로 들어온 입력을 받지 않는다."""

        self._paused = not self._paused
        self._refresh_status()

    # --- Outlook 미러 ---------------------------------------------------

    def _on_mirror_event(self, kind: str, payload: object) -> None:
        """Runs on the sync thread; hand the result to the Tk loop."""

        self._events.put((kind, payload))

    def _is_idle(self) -> bool:
        return not self._queue_count() and (
            self.review is None or not self.review.winfo_exists()
        )

    def _on_mirror_synced(self, payload: object) -> None:
        self.mirror_var.set(f"동기화 {payload}")
        self.mirror_detail = f"Outlook {payload}건 동기화됨"
        if self._is_idle():
            self._refresh_status()

    def _on_mirror_error(self, error: str) -> None:
        # 동기화 실패는 아래 줄에만 남긴다. 큰 상태는 사용자가 할 일을 가리켜야 한다.
        self.mirror_var.set("동기화 실패")
        self.mirror_detail = f"Outlook 동기화 실패 · 등록 시 직접 확인 ({error[:40]})"
        if self._is_idle():
            self._refresh_status()

    def _update_queue_status(self) -> None:
        pending = self._queue_count()
        review = self.review
        if review is not None and review.winfo_exists():
            review.set_queue_status(pending)
            if self.review is None:  # 마지막 카드까지 처리되면 창이 닫힌다.
                return
        self._refresh_status()

    def _refresh_status(self) -> None:
        """창이 좁으므로 짧은 말만 넣는다. 긴 설명은 `⋯` 메뉴에 있다."""

        pending = self._queue_count()
        if self._paused:
            self._set_cat("idle")
            self.status_var.set("일시중지")
            self.detail_var.set("일시중지됨 · 트레이나 ⋯ 메뉴에서 재개")
        elif self.review is not None and self.review.winfo_exists():
            self._set_cat("eating" if pending else "done")
            self.status_var.set(f"검토 중\n분석 {pending}" if pending else "검토 중")
            self.detail_var.set(
                f"검토 중 · 분석 중인 메시지 {pending}건" if pending else "검토 중입니다"
            )
        elif self._analysis_in_flight:
            self._set_cat("eating")
            self.status_var.set(f"처리 중\n{self._analysis_in_flight}건")
            self.detail_var.set(f"처리 중입니다 · {self._analysis_in_flight}건")
        elif self._ready_analyses:
            self._set_cat("done")
            self.status_var.set(f"대기\n{len(self._ready_analyses)}건")
            self.detail_var.set(f"대기 중인 일정 {len(self._ready_analyses)}건")
        else:
            self._set_cat("idle")
            self.status_var.set("대기 중")
            self.detail_var.set(f"파일을 놓아 주세요 · {HOTKEY_LABEL}")
        self._sync_tray_tooltip()

    def _set_cat(self, state: str) -> None:
        """표정만 바꾼다. 문구는 status_var 한 곳에서만 관리한다."""

        self.face_var.set(FACE.get(state, FACE["idle"]))

    def _drop_hover(self, _event: tk.Event) -> None:
        hover = "#dbeefa"
        self.drop_zone.configure(bg=hover)
        for widget in (self.face_label, self.state_label):
            widget.configure(bg=hover)

    def _drop_leave(self, _event: tk.Event) -> None:
        self.drop_zone.configure(bg=theme.ACCENT_SOFT)
        for widget in (self.face_label, self.state_label):
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
        self._enqueue_analysis(raw, source="drop")

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

    def paste_from_clipboard(self, _event: tk.Event | None = None) -> str | None:
        """입력창 없이 Ctrl+V 흐름을 유지한다."""

        try:
            text = str(self.clipboard_get()).strip()
        except tk.TclError:
            text = ""
        if not text:
            self._set_cat("error")
            self.status_var.set("클립보드에 텍스트가 없습니다.")
            self.bell()
            return "break" if _event is not None else None
        self._enqueue_analysis(text, source="clipboard")
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
                elif kind == "mirror_synced":
                    self._on_mirror_synced(payload)
                elif kind == "mirror_error":
                    self._on_mirror_error(str(payload))
                elif kind == "tray_show":
                    self.show_window()
                elif kind == "tray_pause":
                    self.toggle_pause()
                elif kind == "tray_quit":
                    self.quit_app()
                    return
                else:
                    self._enqueue_analysis(
                        str(payload),
                        source="selection" if kind == "selection" else "clipboard",
                    )
        except queue.Empty:
            pass
        finally:
            # 트레이 `종료`로 인터프리터가 이미 사라졌을 수 있다.
            try:
                if self.winfo_exists():
                    self.after(100, self._drain_events)
            except tk.TclError:
                pass

    def _on_files_loaded(self, result: object) -> None:
        if not isinstance(result, FileLoadResult):
            self._set_cat("error")
            self.status_var.set("파일 처리 결과를 읽지 못했습니다.")
            return
        self._enqueue_analysis(result.text, source="file")
        # 대기열 상태가 먼저 찍히므로, 알릴 것이 있을 때만 덮어쓴다.
        if result.warnings:
            self.status_var.set("일부 파일을 읽지 못했어요 · 처리 중")

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
        review = self.review if self.review is not None and self.review.winfo_exists() else None
        if review is not None:
            # 검토하는 동안 분석이 끝난 메시지는 기다리지 않고 같은 창에 붙인다.
            while self._ready_analyses and review.winfo_exists():
                review.present_analysis(self._ready_analyses.popleft())
            self._update_queue_status()
            return
        if self._ready_analyses:
            analysis = self._ready_analyses.popleft()
            review = self._review_window()
            review.present_analysis(analysis)
            self._update_queue_status()
            return
        if self._analysis_errors and not self._analysis_in_flight:
            error = self._analysis_errors.popleft()
            self._set_cat("error")
            self.status_var.set(f"분석 실패: {error[:60]}")
            self.bell()
            return
        self._refresh_status()

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
        # 창이 좁아 짧은 상태만 보이므로, 자세한 문장은 여기 맨 위에 둔다.
        menu.add_command(label=self.detail_var.get(), state="disabled")
        menu.add_command(label=self.mirror_detail, state="disabled")
        mode = "Solar Pro 4 AI" if self.service.has_api_key() else "오프라인 규칙"
        menu.add_command(label=f"분석 방식 · {mode}", state="disabled")
        menu.add_separator()
        menu.add_command(label=f"{HOTKEY_LABEL} 선택 텍스트 가져오기", command=self.capture_now)
        menu.add_command(label="붙여넣기 (Ctrl+V)", command=self.paste_from_clipboard)
        menu.add_command(label="파일 선택", command=self.choose_files)
        menu.add_separator()
        menu.add_command(
            label="재개" if self._paused else "일시중지", command=self.toggle_pause
        )
        menu.add_command(label="Outlook 연결 확인", command=self.test_outlook)
        menu.add_command(label="단축키 안내", command=self._show_shortcuts)
        menu.add_separator()
        menu.add_command(label="종료", command=self.quit_app)
        try:
            menu.tk_popup(self.winfo_rootx() + self.winfo_width() - 150, self.winfo_rooty() + 32)
        finally:
            menu.grab_release()

    def _show_shortcuts(self) -> None:
        messagebox.showinfo(
            "사용 방법",
            f"{HOTKEY_LABEL}: 선택한 텍스트 가져오기\n"
            "Ctrl+V: 클립보드 내용 분석\n"
            "파일 드롭 또는 창 클릭: 문서·이미지 읽기\n\n"
            "창을 닫아도 트레이 아이콘에서 일시중지·종료할 수 있습니다.",
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
        if self.tray is not None:
            self.tray.stop()
        self.service.mirror.stop()
        self._analysis_jobs.put(None)
        self.destroy()


def main() -> None:
    # 창을 만들기 전에 불러야 작업 표시줄이 앱 아이콘을 쓴다.
    icons.claim_taskbar_identity()
    MiniWindow().mainloop()


if __name__ == "__main__":
    main()
