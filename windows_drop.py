"""Small native Windows file-drop bridge for a Tk top-level window."""

from __future__ import annotations

import ctypes
import os
from ctypes import wintypes
from typing import Callable


class WindowsFileDropHandler:
    """Deliver paths dropped on a Tk window without adding a TkDND runtime."""

    WM_DROPFILES = 0x0233
    GWL_WNDPROC = -4

    def __init__(self, widget: object, callback: Callable[[list[str]], None]) -> None:
        self.widget = widget
        self.callback = callback
        self._user32 = None
        self._shell32 = None
        self._hwnd: int | None = None
        self._old_proc: int | None = None
        self._new_proc = None

        if os.name != "nt":
            return
        self._install()

    def _install(self) -> None:
        import tkinter as tk

        if not isinstance(self.widget, tk.Misc):
            raise TypeError("WindowsFileDropHandler에는 Tk 위젯이 필요합니다.")

        hwnd = int(self.widget.winfo_id())
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        shell32 = ctypes.WinDLL("shell32", use_last_error=True)
        shell32.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]
        shell32.DragAcceptFiles.restype = None
        shell32.DragQueryFileW.argtypes = [
            wintypes.HANDLE,
            wintypes.UINT,
            wintypes.LPWSTR,
            wintypes.UINT,
        ]
        shell32.DragQueryFileW.restype = wintypes.UINT
        shell32.DragFinish.argtypes = [wintypes.HANDLE]
        shell32.DragFinish.restype = None
        user32.SetWindowLongPtrW.argtypes = [
            wintypes.HWND,
            ctypes.c_int,
            ctypes.c_void_p,
        ]
        user32.SetWindowLongPtrW.restype = ctypes.c_void_p
        user32.CallWindowProcW.argtypes = [
            ctypes.c_void_p,
            wintypes.HWND,
            wintypes.UINT,
            wintypes.WPARAM,
            wintypes.LPARAM,
        ]
        user32.CallWindowProcW.restype = ctypes.c_ssize_t

        result_type = ctypes.c_ssize_t
        proc_type = ctypes.WINFUNCTYPE(
            result_type,
            wintypes.HWND,
            wintypes.UINT,
            wintypes.WPARAM,
            wintypes.LPARAM,
        )

        self._user32 = user32
        self._shell32 = shell32
        self._hwnd = hwnd
        self._new_proc = proc_type(self._wnd_proc)
        old_proc = user32.SetWindowLongPtrW(hwnd, self.GWL_WNDPROC, self._new_proc)
        if not old_proc:
            self._new_proc = None
            self._hwnd = None
            raise ctypes.WinError(ctypes.get_last_error())
        self._old_proc = int(old_proc)
        shell32.DragAcceptFiles(hwnd, True)

    def _read_paths(self, handle: int) -> list[str]:
        assert self._shell32 is not None
        count = int(self._shell32.DragQueryFileW(handle, 0xFFFFFFFF, None, 0))
        paths: list[str] = []
        for index in range(count):
            length = int(self._shell32.DragQueryFileW(handle, index, None, 0))
            buffer = ctypes.create_unicode_buffer(length + 1)
            self._shell32.DragQueryFileW(handle, index, buffer, length + 1)
            paths.append(buffer.value)
        return paths

    def _wnd_proc(self, hwnd: int, message: int, wparam: int, lparam: int) -> int:
        if message == self.WM_DROPFILES and self._shell32 is not None:
            try:
                paths = self._read_paths(wparam)
            except Exception:
                paths = []
            finally:
                try:
                    self._shell32.DragFinish(wparam)
                except Exception:
                    pass
            if paths:
                # Keep all Tk work on the Tk event loop.  The native callback
                # itself must stay tiny and must never parse a file.
                try:
                    self.widget.after_idle(self.callback, paths)
                except Exception:
                    pass
            return 0

        if self._old_proc and self._user32 is not None:
            return int(self._user32.CallWindowProcW(
                self._old_proc, hwnd, message, wparam, lparam
            ))
        return 0

    def close(self) -> None:
        if self._user32 is None or self._hwnd is None:
            return
        try:
            if self._shell32 is not None:
                self._shell32.DragAcceptFiles(self._hwnd, False)
            if self._old_proc:
                self._user32.SetWindowLongPtrW(
                    self._hwnd, self.GWL_WNDPROC, ctypes.c_void_p(self._old_proc)
                )
        finally:
            self._old_proc = None
            self._new_proc = None
            self._hwnd = None
