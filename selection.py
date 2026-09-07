"""Read the text the user has selected in any Windows application.

The hotkey fires while another program has focus, so there is no API that
returns "the current selection" for arbitrary apps. The portable approach is
to synthesise Ctrl+C, wait for the clipboard to change, then put the previous
clipboard content back.
"""

from __future__ import annotations

import ctypes
import time
from ctypes import wintypes
from dataclasses import dataclass

user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

if ctypes.sizeof(ctypes.c_void_p) == 8:
    ULONG_PTR = ctypes.c_uint64
else:  # pragma: no cover - 32-bit Python is not the target environment
    ULONG_PTR = ctypes.c_uint32

CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002
KEYEVENTF_KEYUP = 0x0002

VK_SHIFT = 0x10
VK_CONTROL = 0x11
VK_MENU = 0x12
VK_LSHIFT = 0xA0
VK_RSHIFT = 0xA1
VK_LCONTROL = 0xA2
VK_RCONTROL = 0xA3
VK_LWIN = 0x5B
VK_RWIN = 0x5C
VK_C = 0x43

# Modifiers that must be physically released before Ctrl+C is synthesised.
# Otherwise the still-held Shift of the hotkey turns it into Ctrl+Shift+C.
_MODIFIER_KEYS = (
    VK_LSHIFT,
    VK_RSHIFT,
    VK_SHIFT,
    VK_LCONTROL,
    VK_RCONTROL,
    VK_CONTROL,
    VK_MENU,
    VK_LWIN,
    VK_RWIN,
)

user32.OpenClipboard.argtypes = [wintypes.HWND]
user32.OpenClipboard.restype = wintypes.BOOL
user32.CloseClipboard.argtypes = []
user32.CloseClipboard.restype = wintypes.BOOL
user32.EmptyClipboard.argtypes = []
user32.EmptyClipboard.restype = wintypes.BOOL
user32.GetClipboardData.argtypes = [wintypes.UINT]
user32.GetClipboardData.restype = wintypes.HANDLE
user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
user32.SetClipboardData.restype = wintypes.HANDLE
user32.IsClipboardFormatAvailable.argtypes = [wintypes.UINT]
user32.IsClipboardFormatAvailable.restype = wintypes.BOOL
user32.GetClipboardSequenceNumber.argtypes = []
user32.GetClipboardSequenceNumber.restype = wintypes.DWORD
user32.GetAsyncKeyState.argtypes = [ctypes.c_int]
user32.GetAsyncKeyState.restype = ctypes.c_short
user32.keybd_event.argtypes = [
    wintypes.BYTE,
    wintypes.BYTE,
    wintypes.DWORD,
    ULONG_PTR,
]
user32.keybd_event.restype = None

kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
kernel32.GlobalLock.restype = wintypes.LPVOID
kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
kernel32.GlobalUnlock.restype = wintypes.BOOL
kernel32.GlobalFree.argtypes = [wintypes.HGLOBAL]
kernel32.GlobalFree.restype = wintypes.HGLOBAL


class ClipboardError(RuntimeError):
    """Raised when the Windows clipboard cannot be read or written."""


@dataclass(frozen=True)
class Selection:
    """Text picked up by the hotkey and where it came from."""

    text: str
    source: str  # "selection" | "clipboard" | "empty"

    @property
    def description(self) -> str:
        if self.source == "selection":
            return "선택한 텍스트"
        if self.source == "clipboard":
            return "클립보드 내용"
        return "가져온 텍스트 없음"


class _ClipboardSession:
    """Open the clipboard with retries; another app may hold it briefly."""

    def __init__(self, attempts: int = 12, delay: float = 0.02) -> None:
        self.attempts = attempts
        self.delay = delay

    def __enter__(self) -> None:
        for _ in range(self.attempts):
            if user32.OpenClipboard(None):
                return None
            time.sleep(self.delay)
        raise ClipboardError("다른 프로그램이 클립보드를 사용 중이라 열 수 없습니다.")

    def __exit__(self, *_exc: object) -> None:
        user32.CloseClipboard()


def read_clipboard_text() -> str:
    """Return the clipboard's Unicode text, or "" when it holds none."""

    try:
        with _ClipboardSession():
            if not user32.IsClipboardFormatAvailable(CF_UNICODETEXT):
                return ""
            handle = user32.GetClipboardData(CF_UNICODETEXT)
            if not handle:
                return ""
            pointer = kernel32.GlobalLock(handle)
            if not pointer:
                return ""
            try:
                return ctypes.c_wchar_p(pointer).value or ""
            finally:
                kernel32.GlobalUnlock(handle)
    except ClipboardError:
        return ""


def write_clipboard_text(text: str) -> None:
    """Replace the clipboard with `text`."""

    data = ctypes.create_unicode_buffer(text)
    size = ctypes.sizeof(data)
    handle = kernel32.GlobalAlloc(GMEM_MOVEABLE, size)
    if not handle:
        raise ClipboardError("클립보드 메모리를 확보하지 못했습니다.")
    pointer = kernel32.GlobalLock(handle)
    if not pointer:
        kernel32.GlobalFree(handle)
        raise ClipboardError("클립보드 메모리를 잠그지 못했습니다.")
    try:
        ctypes.memmove(pointer, ctypes.byref(data), size)
    finally:
        kernel32.GlobalUnlock(handle)

    try:
        with _ClipboardSession():
            user32.EmptyClipboard()
            if not user32.SetClipboardData(CF_UNICODETEXT, handle):
                raise ClipboardError("클립보드에 쓰지 못했습니다.")
    except ClipboardError:
        kernel32.GlobalFree(handle)
        raise
    # Windows owns the handle once SetClipboardData succeeds.


def _release_held_keys(extra: tuple[int, ...] = ()) -> None:
    for key in (*_MODIFIER_KEYS, *extra):
        if user32.GetAsyncKeyState(key) & 0x8000:
            user32.keybd_event(key, 0, KEYEVENTF_KEYUP, 0)
    time.sleep(0.05)


def _send_ctrl_c() -> None:
    user32.keybd_event(VK_CONTROL, 0, 0, 0)
    user32.keybd_event(VK_C, 0, 0, 0)
    user32.keybd_event(VK_C, 0, KEYEVENTF_KEYUP, 0)
    user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)


def capture_selection(
    *,
    timeout: float = 1.2,
    restore_clipboard: bool = True,
    release_keys: tuple[int, ...] = (),
) -> Selection:
    """Copy the focused app's selection, falling back to the clipboard.

    `release_keys` should list the non-modifier keys of the hotkey (for
    Ctrl+Shift+X that is the X) so they are not still held down when Ctrl+C
    is synthesised.
    """

    previous = read_clipboard_text()
    before = user32.GetClipboardSequenceNumber()

    _release_held_keys(release_keys)
    _send_ctrl_c()

    copied = ""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        time.sleep(0.03)
        if user32.GetClipboardSequenceNumber() != before:
            copied = read_clipboard_text().strip()
            if copied:
                break

    if copied:
        if restore_clipboard and previous and previous != copied:
            try:
                write_clipboard_text(previous)
            except ClipboardError:
                pass  # Leaving the copied text behind is harmless.
        return Selection(text=copied, source="selection")

    previous = previous.strip()
    if previous:
        return Selection(text=previous, source="clipboard")
    return Selection(text="", source="empty")
