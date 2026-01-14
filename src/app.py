import logging
import os
import sys
import tempfile
import atexit
from typing import Optional

_single_instance_handle = None
_single_instance_lock_path: Optional[str] = None

try:
    from src.ui.main_window import MainWindow
except ModuleNotFoundError:
    sys.path.append(os.path.dirname(os.path.dirname(__file__)))
    from src.ui.main_window import MainWindow


def _show_startup_error(title: str, message: str) -> None:
    if sys.platform == "win32":
        try:
            import ctypes

            ctypes.windll.user32.MessageBoxW(0, message, title, 0x10)
            return
        except Exception:
            pass
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(title, message)
        root.destroy()
    except Exception:
        print(f"{title}: {message}", file=sys.stderr)


def _acquire_single_instance_lock() -> bool:
    global _single_instance_handle, _single_instance_lock_path

    if sys.platform == "win32":
        import ctypes

        kernel32 = ctypes.windll.kernel32
        handle = kernel32.CreateMutexW(None, False, "Global\\latex2word_single_instance")
        if not handle:
            return True

        already_exists = int(kernel32.GetLastError()) == 183
        if already_exists:
            try:
                kernel32.CloseHandle(handle)
            except Exception:
                pass
            return False

        _single_instance_handle = handle
        atexit.register(lambda h=handle: kernel32.CloseHandle(h))
        return True

    lock_path = os.path.join(tempfile.gettempdir(), "latex2word.lock")
    try:
        fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        return False

    try:
        os.write(fd, str(os.getpid()).encode("utf-8"))
    finally:
        os.close(fd)

    _single_instance_lock_path = lock_path

    def _cleanup() -> None:
        path = _single_instance_lock_path
        if not path:
            return
        try:
            os.remove(path)
        except OSError:
            pass

    atexit.register(_cleanup)
    return True


def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )
    if not _acquire_single_instance_lock():
        _show_startup_error(
            "latex2word 已在运行",
            "检测到已有 latex2word 实例正在运行。\n请不要重复启动，以避免多个实例产生竞态问题。",
        )
        return 1
    w = MainWindow()
    w.run()
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
