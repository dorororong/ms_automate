"""Windows 알림 영역(트레이) 아이콘.

작은 창에서 버튼을 걷어내면서, 앱을 잠시 멈추거나 끄는 길이 창 안에만 있으면
안 됩니다. 트레이 아이콘은 그 두 가지를 창과 무관하게 제공합니다.

`hotkey.py`와 같은 구조입니다. 트레이 아이콘은 자신을 만든 스레드의 메시지 루프로
알림을 받으므로, 전용 스레드에서 숨은 창을 만들고 그 스레드가 메시지를 폅니다.
메뉴를 누르면 콜백이 그 스레드에서 실행되므로, 호출자는 Tk 스레드로 넘겨야 합니다.
"""

from __future__ import annotations

import threading
from typing import Callable

import icons

try:
    import win32api
    import win32con
    import win32gui
except ImportError:  # pragma: no cover - pywin32 미설치 환경
    win32api = win32con = win32gui = None


class TrayError(RuntimeError):
    """트레이 아이콘을 만들지 못했습니다."""


WM_TRAY = 0x0400 + 20  # WM_USER + 20
ID_SHOW = 1023
ID_PAUSE = 1024
ID_QUIT = 1025


class TrayIcon:
    """전용 스레드에서 숨은 창과 트레이 아이콘을 소유한다."""

    def __init__(
        self,
        title: str,
        *,
        on_show: Callable[[], None],
        on_toggle_pause: Callable[[], None],
        on_quit: Callable[[], None],
        pause_label: Callable[[], str],
    ) -> None:
        self.title = title
        self.on_show = on_show
        self.on_toggle_pause = on_toggle_pause
        self.on_quit = on_quit
        self.pause_label = pause_label
        self._hwnd: int | None = None
        self._thread: threading.Thread | None = None
        self._ready = threading.Event()
        self._error: str | None = None
        self._lock = threading.Lock()
        self._tooltip = title
        self._icon: int | None = None

    # --- 수명 ------------------------------------------------------------

    def start(self) -> None:
        if win32gui is None:
            raise TrayError("pywin32가 없어 트레이 아이콘을 만들 수 없습니다.")
        if self._thread is not None and self._thread.is_alive():
            return
        self._ready.clear()
        self._thread = threading.Thread(
            target=self._run, name="tray-icon", daemon=True
        )
        self._thread.start()
        # 아이콘 등록 실패를 호출자가 바로 알 수 있어야 한다.
        self._ready.wait(timeout=5.0)
        if self._error:
            raise TrayError(self._error)

    def stop(self) -> None:
        hwnd = self._hwnd
        if hwnd is None or win32gui is None:
            return
        try:
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        except Exception:
            pass

    def set_tooltip(self, text: str) -> None:
        """툴팁은 트레이 창 스레드에서만 바꾼다."""

        with self._lock:
            self._tooltip = text[:127] or self.title
        hwnd = self._hwnd
        if hwnd is None or win32gui is None:
            return
        try:
            win32gui.PostMessage(hwnd, WM_TRAY + 1, 0, 0)
        except Exception:
            pass

    # --- 트레이 스레드 ----------------------------------------------------

    def _run(self) -> None:
        try:
            self._create_window()
            self._add_icon()
        except Exception as exc:
            self._error = str(exc)
            self._ready.set()
            return
        self._error = None
        self._ready.set()
        try:
            win32gui.PumpMessages()
        finally:
            self._remove_icon()

    def _create_window(self) -> None:
        instance = win32api.GetModuleHandle(None)
        window_class = win32gui.WNDCLASS()
        window_class.hInstance = instance
        window_class.lpszClassName = "MsAutomateTrayWindow"
        window_class.lpfnWndProc = {
            WM_TRAY: self._on_tray_message,
            WM_TRAY + 1: self._on_tooltip_message,
            win32con.WM_COMMAND: self._on_command,
            win32con.WM_CLOSE: self._on_close,
            win32con.WM_DESTROY: self._on_destroy,
        }
        try:
            atom = win32gui.RegisterClass(window_class)
        except Exception:
            # 같은 프로세스에서 다시 시작하면 클래스가 이미 등록돼 있다.
            atom = window_class.lpszClassName
        self._hwnd = win32gui.CreateWindow(
            atom, self.title, win32con.WS_OVERLAPPED,
            0, 0, 0, 0, 0, 0, instance, None,
        )
        win32gui.UpdateWindow(self._hwnd)

    def _load_icon(self) -> int:
        """앱 아이콘을 알림 영역 크기로 읽는다. 실패하면 기본 아이콘."""

        path = icons.ico_path()
        if path is not None:
            try:
                return win32gui.LoadImage(
                    0,
                    str(path),
                    win32con.IMAGE_ICON,
                    win32api.GetSystemMetrics(win32con.SM_CXSMICON),
                    win32api.GetSystemMetrics(win32con.SM_CYSMICON),
                    win32con.LR_LOADFROMFILE,
                )
            except Exception:
                pass
        return win32gui.LoadIcon(0, win32con.IDI_APPLICATION)

    def _notify_data(self) -> tuple:
        with self._lock:
            tooltip = self._tooltip
        if self._icon is None:
            self._icon = self._load_icon()
        flags = win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP
        return (self._hwnd, 0, flags, WM_TRAY, self._icon, tooltip)

    def _add_icon(self) -> None:
        win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, self._notify_data())

    def _remove_icon(self) -> None:
        if self._hwnd is None:
            return
        try:
            win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, (self._hwnd, 0))
        except Exception:
            pass
        self._hwnd = None

    # --- 메시지 처리 ------------------------------------------------------

    def _on_tooltip_message(self, hwnd: int, msg: int, wparam: int, lparam: int) -> int:
        try:
            win32gui.Shell_NotifyIcon(win32gui.NIM_MODIFY, self._notify_data())
        except Exception:
            pass
        return 0

    def _on_tray_message(self, hwnd: int, msg: int, wparam: int, lparam: int) -> int:
        if lparam == win32con.WM_LBUTTONDBLCLK or lparam == win32con.WM_LBUTTONUP:
            self._safe(self.on_show)
        elif lparam == win32con.WM_RBUTTONUP:
            self._show_menu(hwnd)
        return 0

    def _show_menu(self, hwnd: int) -> None:
        menu = win32gui.CreatePopupMenu()
        win32gui.AppendMenu(menu, win32con.MF_STRING, ID_SHOW, "창 보이기")
        win32gui.AppendMenu(menu, win32con.MF_STRING, ID_PAUSE, self._pause_text())
        win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "")
        win32gui.AppendMenu(menu, win32con.MF_STRING, ID_QUIT, "종료")
        position = win32gui.GetCursorPos()
        # 메뉴 밖을 눌렀을 때 닫히게 하려면 포그라운드를 잡아야 한다.
        win32gui.SetForegroundWindow(hwnd)
        win32gui.TrackPopupMenu(
            menu, win32con.TPM_LEFTALIGN | win32con.TPM_RIGHTBUTTON,
            position[0], position[1], 0, hwnd, None,
        )
        win32gui.PostMessage(hwnd, win32con.WM_NULL, 0, 0)
        win32gui.DestroyMenu(menu)

    def _pause_text(self) -> str:
        try:
            return self.pause_label()
        except Exception:
            return "일시중지"

    def _on_command(self, hwnd: int, msg: int, wparam: int, lparam: int) -> int:
        command = win32api.LOWORD(wparam)
        if command == ID_SHOW:
            self._safe(self.on_show)
        elif command == ID_PAUSE:
            self._safe(self.on_toggle_pause)
        elif command == ID_QUIT:
            self._safe(self.on_quit)
        return 0

    def _on_close(self, hwnd: int, msg: int, wparam: int, lparam: int) -> int:
        win32gui.DestroyWindow(hwnd)
        return 0

    def _on_destroy(self, hwnd: int, msg: int, wparam: int, lparam: int) -> int:
        self._remove_icon()
        win32gui.PostQuitMessage(0)
        return 0

    @staticmethod
    def _safe(callback: Callable[[], None]) -> None:
        """콜백 실패가 트레이 스레드를 죽이지 않게 한다."""

        try:
            callback()
        except Exception:
            pass
