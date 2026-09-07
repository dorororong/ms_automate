"""System-wide hotkey registration for Windows.

Uses RegisterHotKey on a dedicated thread with its own message loop, so the
shortcut works while any other application has focus. No extra dependency is
needed beyond what Python already ships (ctypes).
"""

from __future__ import annotations

import ctypes
import threading
from ctypes import wintypes
from typing import Callable

user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008
MOD_NOREPEAT = 0x4000

WM_HOTKEY = 0x0312
WM_QUIT = 0x0012

ERROR_HOTKEY_ALREADY_REGISTERED = 1409

user32.RegisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int, wintypes.UINT, wintypes.UINT]
user32.RegisterHotKey.restype = wintypes.BOOL
user32.UnregisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int]
user32.UnregisterHotKey.restype = wintypes.BOOL
user32.GetMessageW.argtypes = [ctypes.POINTER(wintypes.MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT]
user32.GetMessageW.restype = ctypes.c_int
user32.PostThreadMessageW.argtypes = [wintypes.DWORD, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.PostThreadMessageW.restype = wintypes.BOOL
kernel32.GetCurrentThreadId.argtypes = []
kernel32.GetCurrentThreadId.restype = wintypes.DWORD


class HotkeyError(RuntimeError):
    """Raised when the shortcut cannot be registered with Windows."""


class HotkeyListener:
    """Call `callback` whenever the registered shortcut is pressed.

    The callback runs on the listener thread, not the Tk thread, so it must
    only hand the event over (a queue) and return quickly.
    """

    HOTKEY_ID = 0xA51

    def __init__(
        self,
        callback: Callable[[], None],
        *,
        modifiers: int = MOD_CONTROL | MOD_SHIFT,
        key_code: int = 0x58,  # 'X'
        label: str = "Ctrl+Shift+X",
    ) -> None:
        self.callback = callback
        self.modifiers = modifiers | MOD_NOREPEAT
        self.key_code = key_code
        self.label = label
        self._thread: threading.Thread | None = None
        self._thread_id: int = 0
        self._ready = threading.Event()
        self._error: str | None = None

    def start(self) -> None:
        if self._thread is not None:
            return
        self._ready.clear()
        self._error = None
        thread = threading.Thread(target=self._run, name="global-hotkey", daemon=True)
        self._thread = thread
        thread.start()
        self._ready.wait(timeout=3.0)
        if self._error:
            self._thread = None
            raise HotkeyError(self._error)

    def stop(self) -> None:
        if self._thread is None:
            return
        if self._thread_id:
            user32.PostThreadMessageW(self._thread_id, WM_QUIT, 0, 0)
        self._thread.join(timeout=2.0)
        self._thread = None
        self._thread_id = 0

    def _run(self) -> None:
        self._thread_id = kernel32.GetCurrentThreadId()
        if not user32.RegisterHotKey(None, self.HOTKEY_ID, self.modifiers, self.key_code):
            code = ctypes.get_last_error()
            if code == ERROR_HOTKEY_ALREADY_REGISTERED:
                self._error = (
                    f"{self.label}은 다른 프로그램이 이미 사용 중입니다. "
                    "그 프로그램을 종료하거나 단축키를 바꾸세요."
                )
            else:
                self._error = f"{self.label} 등록 실패 (Windows 오류 {code})."
            self._ready.set()
            return

        self._ready.set()
        message = wintypes.MSG()
        try:
            while True:
                result = user32.GetMessageW(ctypes.byref(message), None, 0, 0)
                if result in (0, -1):
                    break
                if message.message == WM_HOTKEY:
                    try:
                        self.callback()
                    except Exception:
                        # A UI failure must not kill the listener thread.
                        pass
        finally:
            user32.UnregisterHotKey(None, self.HOTKEY_ID)
