from __future__ import annotations

import atexit
import ctypes
import threading
from ctypes import wintypes


class TrayIcon:
    def __init__(self, on_show, on_toggle_auto_paste, on_exit, get_auto_paste_state, *, icon_path: str | None = None) -> None:
        self._on_show = on_show
        self._on_toggle_auto_paste = on_toggle_auto_paste
        self._on_exit = on_exit
        self._get_auto_paste_state = get_auto_paste_state
        self._icon_path = icon_path
        self._thread: threading.Thread | None = None
        self._hwnd: int = 0
        self._hicon: int = 0
        self._owns_hicon = False
        self._stop_evt = threading.Event()

    def start(self) -> None:
        if self._thread is not None:
            return
        t = threading.Thread(target=self._run, daemon=True)
        self._thread = t
        t.start()
        atexit.register(self.stop)

    def stop(self) -> None:
        self._stop_evt.set()
        hwnd = self._hwnd
        if hwnd:
            user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)

    def _run(self) -> None:
        self._ensure_winapi()
        hinst = kernel32.GetModuleHandleW(None)
        class_name = "latex2word_tray"

        wndproc = WNDPROC(self._wnd_proc)
        wc = WNDCLASSW()
        wc.lpfnWndProc = wndproc
        wc.hInstance = hinst
        wc.lpszClassName = class_name
        user32.RegisterClassW(ctypes.byref(wc))

        hwnd = user32.CreateWindowExW(0, class_name, class_name, 0, 0, 0, 0, 0, 0, 0, hinst, None)
        self._hwnd = hwnd

        self._add_icon(hwnd)

        msg = MSG()
        while not self._stop_evt.is_set() and user32.GetMessageW(ctypes.byref(msg), 0, 0, 0) != 0:
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))

        self._remove_icon(hwnd)
        try:
            user32.DestroyWindow(hwnd)
        except Exception:
            pass
        self._hwnd = 0

    def _add_icon(self, hwnd: int) -> None:
        nid = NOTIFYICONDATAW()
        nid.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
        nid.hWnd = hwnd
        nid.uID = 1
        nid.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP
        nid.uCallbackMessage = TRAY_CALLBACK_MSG
        hicon = 0
        icon_path = self._icon_path
        if icon_path:
            hicon = self._create_hicon_from_ico(icon_path)
            self._owns_hicon = bool(hicon)
        if not hicon:
            hicon = user32.LoadIconW(None, ctypes.cast(ctypes.c_void_p(IDI_APPLICATION), wintypes.LPCWSTR))
            self._owns_hicon = False
        self._hicon = int(hicon)
        nid.hIcon = hicon
        nid.szTip = "latex2word"
        shell32.Shell_NotifyIconW(NIM_ADD, ctypes.byref(nid))

    def _remove_icon(self, hwnd: int) -> None:
        nid = NOTIFYICONDATAW()
        nid.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
        nid.hWnd = hwnd
        nid.uID = 1
        shell32.Shell_NotifyIconW(NIM_DELETE, ctypes.byref(nid))
        if self._owns_hicon and self._hicon:
            try:
                user32.DestroyIcon(wintypes.HICON(self._hicon))
            except Exception:
                pass
        self._hicon = 0
        self._owns_hicon = False

    def _create_hicon_from_ico(self, ico_path: str) -> int:
        try:
            hicon = user32.LoadImageW(
                0,
                ctypes.c_wchar_p(ico_path),
                IMAGE_ICON,
                0,
                0,
                LR_LOADFROMFILE | LR_DEFAULTSIZE,
            )
            return int(hicon)
        except Exception:
            return 0

    def _show_menu(self, hwnd: int) -> None:
        menu = user32.CreatePopupMenu()
        user32.AppendMenuW(menu, MF_STRING, 1001, "显示主页面")
        state = self._get_auto_paste_state()
        user32.AppendMenuW(menu, MF_STRING, 1002, "无感粘贴：开" if state else "无感粘贴：关")
        user32.AppendMenuW(menu, MF_SEPARATOR, 0, None)
        user32.AppendMenuW(menu, MF_STRING, 1003, "退出程序")

        pt = POINT()
        user32.GetCursorPos(ctypes.byref(pt))
        user32.SetForegroundWindow(hwnd)
        user32.TrackPopupMenu(menu, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, None)
        user32.DestroyMenu(menu)

    def _wnd_proc(self, hwnd, msg, wparam, lparam):
        if msg == TRAY_CALLBACK_MSG:
            if lparam == WM_RBUTTONUP:
                self._show_menu(hwnd)
                return 0
        if msg == WM_COMMAND:
            cmd_id = wparam & 0xFFFF
            if cmd_id == 1001:
                self._on_show()
                return 0
            if cmd_id == 1002:
                self._on_toggle_auto_paste()
                return 0
            if cmd_id == 1003:
                self._on_exit()
                return 0
        if msg == WM_CLOSE:
            user32.PostQuitMessage(0)
            return 0
        return user32.DefWindowProcW(hwnd, msg, wparam, lparam)

    def _ensure_winapi(self) -> None:
        global user32, kernel32, shell32
        global WNDCLASSW, WNDPROC, MSG, POINT, NOTIFYICONDATAW
        global WM_COMMAND, WM_CLOSE, WM_RBUTTONUP, TRAY_CALLBACK_MSG
        global NIF_MESSAGE, NIF_ICON, NIF_TIP, NIM_ADD, NIM_DELETE, IDI_APPLICATION
        global IMAGE_ICON, LR_LOADFROMFILE, LR_DEFAULTSIZE
        global MF_STRING, MF_SEPARATOR, TPM_RIGHTBUTTON

        user32 = ctypes.WinDLL("user32", use_last_error=True)
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        shell32 = ctypes.WinDLL("shell32", use_last_error=True)

        WM_COMMAND = 0x0111
        WM_CLOSE = 0x0010
        WM_RBUTTONUP = 0x0205
        TRAY_CALLBACK_MSG = 0x8001

        NIF_MESSAGE = 0x00000001
        NIF_ICON = 0x00000002
        NIF_TIP = 0x00000004
        NIM_ADD = 0x00000000
        NIM_DELETE = 0x00000002
        IDI_APPLICATION = 32512
        IMAGE_ICON = 1
        LR_LOADFROMFILE = 0x00000010
        LR_DEFAULTSIZE = 0x00000040

        MF_STRING = 0x00000000
        MF_SEPARATOR = 0x00000800
        TPM_RIGHTBUTTON = 0x0002

        LRESULT = getattr(wintypes, "LRESULT", ctypes.c_ssize_t)
        UINT_PTR = getattr(wintypes, "UINT_PTR", ctypes.c_size_t)
        LPCRECT = getattr(wintypes, "LPCRECT", getattr(wintypes, "LPRECT", ctypes.c_void_p))
        HCURSOR = getattr(wintypes, "HCURSOR", wintypes.HANDLE)
        HBRUSH = getattr(wintypes, "HBRUSH", wintypes.HANDLE)

        WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)

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

        class POINT(ctypes.Structure):
            _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

        class MSG(ctypes.Structure):
            _fields_ = [
                ("hwnd", wintypes.HWND),
                ("message", wintypes.UINT),
                ("wParam", wintypes.WPARAM),
                ("lParam", wintypes.LPARAM),
                ("time", wintypes.DWORD),
                ("pt", POINT),
            ]

        class NOTIFYICONDATAW(ctypes.Structure):
            _fields_ = [
                ("cbSize", wintypes.DWORD),
                ("hWnd", wintypes.HWND),
                ("uID", wintypes.UINT),
                ("uFlags", wintypes.UINT),
                ("uCallbackMessage", wintypes.UINT),
                ("hIcon", wintypes.HICON),
                ("szTip", wintypes.WCHAR * 128),
                ("dwState", wintypes.DWORD),
                ("dwStateMask", wintypes.DWORD),
                ("szInfo", wintypes.WCHAR * 256),
                ("uTimeoutOrVersion", wintypes.UINT),
                ("szInfoTitle", wintypes.WCHAR * 64),
                ("dwInfoFlags", wintypes.DWORD),
                ("guidItem", ctypes.c_byte * 16),
                ("hBalloonIcon", wintypes.HICON),
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
        user32.LoadImageW.argtypes = [
            wintypes.HINSTANCE,
            wintypes.LPCWSTR,
            wintypes.UINT,
            ctypes.c_int,
            ctypes.c_int,
            wintypes.UINT,
        ]
        user32.LoadImageW.restype = wintypes.HANDLE
        user32.LoadIconW.argtypes = [wintypes.HINSTANCE, wintypes.LPCWSTR]
        user32.LoadIconW.restype = wintypes.HICON
        user32.DestroyIcon.argtypes = [wintypes.HICON]
        user32.DestroyIcon.restype = wintypes.BOOL
        user32.CreatePopupMenu.argtypes = []
        user32.CreatePopupMenu.restype = wintypes.HMENU
        user32.AppendMenuW.argtypes = [wintypes.HMENU, wintypes.UINT, UINT_PTR, wintypes.LPCWSTR]
        user32.AppendMenuW.restype = wintypes.BOOL
        user32.GetCursorPos.argtypes = [ctypes.POINTER(POINT)]
        user32.GetCursorPos.restype = wintypes.BOOL
        user32.SetForegroundWindow.argtypes = [wintypes.HWND]
        user32.SetForegroundWindow.restype = wintypes.BOOL
        user32.TrackPopupMenu.argtypes = [
            wintypes.HMENU,
            wintypes.UINT,
            ctypes.c_int,
            ctypes.c_int,
            ctypes.c_int,
            wintypes.HWND,
            LPCRECT,
        ]
        user32.TrackPopupMenu.restype = wintypes.BOOL
        user32.DestroyMenu.argtypes = [wintypes.HMENU]
        user32.DestroyMenu.restype = wintypes.BOOL

        kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
        kernel32.GetModuleHandleW.restype = wintypes.HMODULE

        shell32.Shell_NotifyIconW.argtypes = [wintypes.DWORD, ctypes.POINTER(NOTIFYICONDATAW)]
        shell32.Shell_NotifyIconW.restype = wintypes.BOOL
