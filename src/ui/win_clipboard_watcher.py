from __future__ import annotations

import ctypes
import logging
import threading
from collections.abc import Callable
from ctypes import wintypes

logger = logging.getLogger(__name__)


class WinClipboardWatcher:
    def __init__(self, *, on_update: Callable[[], None]) -> None:
        self._on_update = on_update
        self._thread: threading.Thread | None = None
        self._hwnd: int = 0
        self._stop_evt = threading.Event()

    def start(self) -> None:
        t0 = self._thread
        if t0 is not None and t0.is_alive():
            return
        self._thread = None
        self._stop_evt.clear()
        logger.info("clipboard watcher start")
        t = threading.Thread(target=self._run, daemon=True)
        self._thread = t
        t.start()

    def stop(self) -> None:
        self._stop_evt.set()
        logger.info("clipboard watcher stop requested")
        hwnd = self._hwnd
        if hwnd:
            user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)

    def _run(self) -> None:
        self._ensure_winapi()
        hwnd = 0
        try:
            hinst = kernel32.GetModuleHandleW(None)
            class_name = "latex_word_clipboard_listener"

            wndproc = WNDPROC(self._wnd_proc)
            wc = WNDCLASSW()
            wc.lpfnWndProc = wndproc
            wc.hInstance = hinst
            wc.lpszClassName = class_name

            if not user32.RegisterClassW(ctypes.byref(wc)):
                if ctypes.get_last_error() != 1410:
                    logger.warning("RegisterClassW failed err=%s", ctypes.get_last_error())
                    return

            hwnd = user32.CreateWindowExW(0, class_name, class_name, 0, 0, 0, 0, 0, 0, 0, hinst, None)
            if not hwnd:
                logger.warning("CreateWindowExW failed err=%s", ctypes.get_last_error())
                return
            self._hwnd = hwnd

            if not user32.AddClipboardFormatListener(hwnd):
                logger.warning("AddClipboardFormatListener failed err=%s", ctypes.get_last_error())
                return

            msg = MSG()
            while not self._stop_evt.is_set() and user32.GetMessageW(ctypes.byref(msg), 0, 0, 0) != 0:
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
        finally:
            if hwnd:
                try:
                    user32.RemoveClipboardFormatListener(hwnd)
                except Exception:
                    pass
                try:
                    user32.DestroyWindow(hwnd)
                except Exception:
                    pass
            self._hwnd = 0
            self._thread = None
            logger.info("clipboard watcher stopped")

    def _wnd_proc(self, hwnd, msg, wparam, lparam):
        if msg == WM_CLIPBOARDUPDATE:
            self._on_update()
            return 0
        if msg == WM_CLOSE:
            user32.PostQuitMessage(0)
            return 0
        return user32.DefWindowProcW(hwnd, msg, wparam, lparam)

    def _ensure_winapi(self) -> None:
        global user32, kernel32
        global WNDCLASSW, WNDPROC, MSG
        global WM_CLOSE, WM_CLIPBOARDUPDATE

        user32 = ctypes.WinDLL("user32", use_last_error=True)
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

        WM_CLOSE = 0x0010
        WM_CLIPBOARDUPDATE = 0x031D

        LRESULT = getattr(wintypes, "LRESULT", ctypes.c_ssize_t)
        WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
        HCURSOR = getattr(wintypes, "HCURSOR", wintypes.HANDLE)
        HBRUSH = getattr(wintypes, "HBRUSH", wintypes.HANDLE)
        POINT = getattr(wintypes, "POINT", None)
        if POINT is None:
            class POINT(ctypes.Structure):
                _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

        class WNDCLASSW(ctypes.Structure):
            _fields_ = [
                ("style", wintypes.UINT),
                ("lpfnWndProc", WNDPROC),
                ("cbClsExtra", ctypes.c_int),
                ("cbWndExtra", ctypes.c_int),
                ("hInstance", wintypes.HINSTANCE),
                ("hIcon", wintypes.HICON),
                ("hCursor", HCURSOR),
                ("hbrBackground", HBRUSH),
                ("lpszMenuName", wintypes.LPCWSTR),
                ("lpszClassName", wintypes.LPCWSTR),
            ]

        class MSG(ctypes.Structure):
            _fields_ = [
                ("hwnd", wintypes.HWND),
                ("message", wintypes.UINT),
                ("wParam", wintypes.WPARAM),
                ("lParam", wintypes.LPARAM),
                ("time", wintypes.DWORD),
                ("pt", POINT),
            ]

        user32.RegisterClassW.argtypes = [ctypes.POINTER(WNDCLASSW)]
        user32.RegisterClassW.restype = wintypes.ATOM
        user32.CreateWindowExW.argtypes = [
            wintypes.DWORD,
            wintypes.LPCWSTR,
            wintypes.LPCWSTR,
            wintypes.DWORD,
            ctypes.c_int,
            ctypes.c_int,
            ctypes.c_int,
            ctypes.c_int,
            wintypes.HWND,
            wintypes.HMENU,
            wintypes.HINSTANCE,
            wintypes.LPVOID,
        ]
        user32.CreateWindowExW.restype = wintypes.HWND

        user32.DefWindowProcW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
        user32.DefWindowProcW.restype = LRESULT
        user32.GetMessageW.argtypes = [ctypes.POINTER(MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT]
        user32.GetMessageW.restype = wintypes.BOOL
        user32.TranslateMessage.argtypes = [ctypes.POINTER(MSG)]
        user32.TranslateMessage.restype = wintypes.BOOL
        user32.DispatchMessageW.argtypes = [ctypes.POINTER(MSG)]
        user32.DispatchMessageW.restype = LRESULT

        user32.PostQuitMessage.argtypes = [ctypes.c_int]
        user32.PostQuitMessage.restype = None
        user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
        user32.PostMessageW.restype = wintypes.BOOL
        user32.DestroyWindow.argtypes = [wintypes.HWND]
        user32.DestroyWindow.restype = wintypes.BOOL

        user32.AddClipboardFormatListener.argtypes = [wintypes.HWND]
        user32.AddClipboardFormatListener.restype = wintypes.BOOL
        user32.RemoveClipboardFormatListener.argtypes = [wintypes.HWND]
        user32.RemoveClipboardFormatListener.restype = wintypes.BOOL

        kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
        kernel32.GetModuleHandleW.restype = wintypes.HMODULE
