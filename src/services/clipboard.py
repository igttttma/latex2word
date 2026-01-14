from __future__ import annotations

import ctypes
from ctypes import wintypes

user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002

user32.OpenClipboard.argtypes = [wintypes.HWND]
user32.OpenClipboard.restype = wintypes.BOOL
user32.CloseClipboard.argtypes = []
user32.CloseClipboard.restype = wintypes.BOOL
user32.EmptyClipboard.argtypes = []
user32.EmptyClipboard.restype = wintypes.BOOL
user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
user32.SetClipboardData.restype = wintypes.HANDLE
user32.IsClipboardFormatAvailable.argtypes = [wintypes.UINT]
user32.IsClipboardFormatAvailable.restype = wintypes.BOOL
user32.GetClipboardData.argtypes = [wintypes.UINT]
user32.GetClipboardData.restype = wintypes.HANDLE

kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
kernel32.GlobalLock.restype = wintypes.LPVOID
kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
kernel32.GlobalUnlock.restype = wintypes.BOOL
kernel32.GlobalFree.argtypes = [wintypes.HGLOBAL]
kernel32.GlobalFree.restype = wintypes.HGLOBAL

def copy_text(text: str) -> None:
    if text is None:
        text = ""
    if not isinstance(text, str):
        text = str(text)

    if not user32.OpenClipboard(None):
        raise OSError(ctypes.get_last_error())
    try:
        if not user32.EmptyClipboard():
            raise OSError(ctypes.get_last_error())

        data = text.encode("utf-16-le") + b"\x00\x00"
        h_global = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(data))
        if not h_global:
            raise MemoryError("GlobalAlloc failed")

        locked = kernel32.GlobalLock(h_global)
        if not locked:
            kernel32.GlobalFree(h_global)
            raise OSError(ctypes.get_last_error())

        try:
            ctypes.memmove(locked, data, len(data))
        finally:
            kernel32.GlobalUnlock(h_global)

        if not user32.SetClipboardData(CF_UNICODETEXT, h_global):
            kernel32.GlobalFree(h_global)
            raise OSError(ctypes.get_last_error())
    finally:
        user32.CloseClipboard()


def get_text() -> str:
    if not user32.OpenClipboard(None):
        raise OSError(ctypes.get_last_error())
    try:
        if not user32.IsClipboardFormatAvailable(CF_UNICODETEXT):
            return ""
        h_data = user32.GetClipboardData(CF_UNICODETEXT)
        if not h_data:
            return ""

        locked = kernel32.GlobalLock(h_data)
        if not locked:
            raise OSError(ctypes.get_last_error())
        try:
            raw = ctypes.wstring_at(locked)
            return raw or ""
        finally:
            kernel32.GlobalUnlock(h_data)
    finally:
        user32.CloseClipboard()
