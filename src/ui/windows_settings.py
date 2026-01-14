from __future__ import annotations

import sys

try:
    import winreg
except Exception:
    winreg = None


def load_settings(*, is_frozen: bool) -> dict[str, object] | None:
    if winreg is None:
        return None
    for key_path in (r"Software\latex2word"):
        try:
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
            close_behavior, _ = winreg.QueryValueEx(key, "close_behavior")
            autostart, _ = winreg.QueryValueEx(key, "autostart")
            auto_paste, _ = winreg.QueryValueEx(key, "auto_paste")
            winreg.CloseKey(key)
            return {
                "close_behavior": close_behavior,
                "autostart": autostart if is_frozen else None,
                "auto_paste": str(auto_paste) == "1",
            }
        except Exception:
            continue
    return None


def persist_settings(*, close_behavior: str, autostart: str, auto_paste: bool, is_frozen: bool) -> None:
    if winreg is None:
        return
    try:
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\latex2word")
        winreg.SetValueEx(key, "close_behavior", 0, winreg.REG_SZ, close_behavior)
        if is_frozen:
            winreg.SetValueEx(key, "autostart", 0, winreg.REG_SZ, autostart)
        winreg.SetValueEx(key, "auto_paste", 0, winreg.REG_SZ, "1" if auto_paste else "0")
        winreg.CloseKey(key)
    except Exception:
        return


def apply_autostart(*, autostart: str, is_frozen: bool) -> None:
    if not is_frozen:
        return
    if winreg is None:
        return
    name = "latex2word"
    run_key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, run_key_path, 0, winreg.KEY_SET_VALUE)
    except Exception:
        return
    try:
        if autostart == "on":
            winreg.SetValueEx(key, name, 0, winreg.REG_SZ, get_startup_command(silent=True))
        else:
            try:
                winreg.DeleteValue(key, name)
            except FileNotFoundError:
                pass
    finally:
        winreg.CloseKey(key)


def get_autostart_state(*, is_frozen: bool) -> str | None:
    if not is_frozen:
        return None
    if winreg is None:
        return None
    name = "latex2word"
    run_key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, run_key_path, 0, winreg.KEY_READ)
    except Exception:
        return None
    try:
        try:
            winreg.QueryValueEx(key, name)
            return "on"
        except FileNotFoundError:
            return "off"
        except OSError:
            return None
    finally:
        winreg.CloseKey(key)


def ensure_autostart_silent(*, is_frozen: bool) -> None:
    if not is_frozen:
        return
    if winreg is None:
        return
    name = "latex2word"
    run_key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            run_key_path,
            0,
            winreg.KEY_QUERY_VALUE | winreg.KEY_SET_VALUE,
        )
    except Exception:
        return
    try:
        try:
            current, _ = winreg.QueryValueEx(key, name)
        except FileNotFoundError:
            return
        desired = get_startup_command(silent=True)
        if str(current) != desired:
            winreg.SetValueEx(key, name, 0, winreg.REG_SZ, desired)
    finally:
        winreg.CloseKey(key)


def get_startup_command(*, silent: bool = False) -> str:
    exe = f"\"{sys.executable}\""
    if silent:
        return f"{exe} --silent"
    return exe
