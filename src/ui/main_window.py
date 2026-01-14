from __future__ import annotations

import sys
import webbrowser
from pathlib import Path
import customtkinter as ctk

from src.converters.latex_to_mathml import convert
from src.services.clipboard import copy_text, get_text
from src.ui.clipboard_auto_paste import ClipboardAutoPaster
from src.ui.tray_icon import TrayIcon
from src.ui import windows_settings


class MainWindow:
    def __init__(self) -> None:
        ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
        ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"
        
        # High DPI awareness is handled automatically by CustomTkinter, 
        # but we can ensure scaling is correct
        ctk.set_widget_scaling(1.0)
        ctk.set_window_scaling(1.0)

        self._is_frozen = bool(getattr(sys, "frozen", False))
        self._start_silent = "--silent" in sys.argv[1:]
        self._centered_once = False
        self._set_windows_app_user_model_id()

        self._root = ctk.CTk()
        self._root.title("LaTeX 转 Word (MathML)")
        self._root.resizable(False, False)
        self._icon_ico_path = self._resource_path("image/icon.ico")
        self._apply_app_icons()
        self._auto_paste_var = ctk.BooleanVar(value=True)
        self._auto_paste_preview_var = ctk.StringVar(value="")
        self._status_var = ctk.StringVar(value="")
        self._close_behavior_var = ctk.StringVar(value="exit")
        self._autostart_var = ctk.StringVar(value="off")
        self._topmost_enabled = False
        self._topmost_button: ctk.CTkButton | None = None
        self._auto_paster = ClipboardAutoPaster(
            root=self._root,
            get_clipboard_text=get_text,
            set_clipboard_text=copy_text,
            convert_latex=convert,
            on_preview=self._auto_paste_preview_var.set,
        )
        self._tray: TrayIcon | None = None

        self._load_settings()
        self._sync_autostart_state()
        self._build_ui()
        self._root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._create_tray()
        self._auto_paster.set_enabled(bool(self._auto_paste_var.get()))
        if self._start_silent:
            self._root.withdraw()

    def _resource_path(self, relative_path: str) -> str:
        base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[2]))
        return str(base_dir / relative_path)

    def _apply_app_icons(self) -> None:
        if sys.platform == "win32":
            try:
                self._root.iconbitmap(self._icon_ico_path)
            except Exception:
                pass

    def _set_windows_app_user_model_id(self) -> None:
        if sys.platform != "win32":
            return
        try:
            import ctypes

            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("latex2word")
        except Exception:
            return

    def run(self) -> None:
        if not self._start_silent:
            self._root.after(0, self._center_window)
        self._root.mainloop()

    def _build_ui(self) -> None:
        # Use a modern font
        font_family = "Microsoft YaHei UI"
        
        # Configure grid weight
        self._root.grid_columnconfigure(0, weight=1)
        self._root.grid_rowconfigure(0, weight=1)

        outer = ctk.CTkFrame(self._root, corner_radius=10)
        outer.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        outer.grid_columnconfigure(0, weight=1)

        # Manual Input Section
        manual = ctk.CTkFrame(outer, corner_radius=6)
        manual.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 0))
        manual.grid_columnconfigure(0, weight=1)
        
        manual_label = ctk.CTkLabel(manual, text="手动输入", font=(font_family, 12, "bold"))
        manual_label.grid(row=0, column=0, sticky="w", padx=10, pady=(5, 0))

        self._text = ctk.CTkTextbox(
            manual,
            width=300,
            height=60,
            font=(font_family, 12),
            wrap="word",
            undo=True
        )
        self._text.grid(row=1, column=0, sticky="ew", padx=10, pady=5)

        manual_bottom = ctk.CTkFrame(manual, fg_color="transparent")
        manual_bottom.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 10))
        manual_bottom.grid_columnconfigure(0, weight=1)

        self._status_label = ctk.CTkLabel(manual_bottom, textvariable=self._status_var, font=(font_family, 12))
        self._status_label.grid(row=0, column=0, sticky="w")

        copy_btn = ctk.CTkButton(
            manual_bottom, 
            text="复制转换后结果", 
            command=self._on_copy_clicked,
            font=(font_family, 12)
        )
        copy_btn.grid(row=0, column=1, sticky="e")

        # Auto Paste Section
        auto = ctk.CTkFrame(outer, corner_radius=6)
        auto.grid(row=1, column=0, sticky="ew", padx=10, pady=(10, 0))
        auto.grid_columnconfigure(0, weight=1)
        
        auto_label = ctk.CTkLabel(auto, text="自动识别", font=(font_family, 12, "bold"))
        auto_label.grid(row=0, column=0, sticky="w", padx=10, pady=(5, 0))

        # Container for description and switch
        auto_content = ctk.CTkFrame(auto, fg_color="transparent")
        auto_content.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 5))
        auto_content.grid_columnconfigure(0, weight=1) # Description takes available space
        auto_content.grid_columnconfigure(1, weight=0) # Switch takes fixed space

        desc = ctk.CTkLabel(
            auto_content,
            text="此模式下，程序将监听剪贴板，识别到公式时自动转换并写回，无需手动操作，直接粘贴即可。",
            wraplength=200,
            justify="left",
            font=(font_family, 12)
        )
        desc.grid(row=0, column=0, sticky="nw")

        self._auto_paste_check = ctk.CTkSwitch(
            auto_content,
            text="开启无感粘贴",
            variable=self._auto_paste_var,
            command=self._on_toggle_auto_paste,
            font=(font_family, 12),
            width=50
        )
        self._auto_paste_check.grid(row=0, column=1, sticky="w", padx=(10, 0))

        preview = ctk.CTkEntry(
            auto, 
            textvariable=self._auto_paste_preview_var, 
            state="readonly", 
            width=300,
            font=(font_family, 12)
        )
        preview.grid(row=2, column=0, sticky="ew", padx=10, pady=10)

        # Settings Section
        settings = ctk.CTkFrame(outer, corner_radius=6)
        settings.grid(row=2, column=0, sticky="ew", padx=10, pady=10)
        settings.grid_columnconfigure(0, weight=1)
        
        settings_label = ctk.CTkLabel(settings, text="程序设置", font=(font_family, 12, "bold"))
        settings_label.grid(row=0, column=0, sticky="w", padx=10, pady=(5, 0))

        row1 = ctk.CTkFrame(settings, fg_color="transparent")
        row1.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
        
        ctk.CTkLabel(row1, text="关闭窗口时", font=(font_family, 12)).pack(side="left")
        ctk.CTkRadioButton(
            row1,
            text="关闭程序",
            value="exit",
            variable=self._close_behavior_var,
            command=self._persist_settings,
            font=(font_family, 12)
        ).pack(side="left", padx=(14, 0))
        ctk.CTkRadioButton(
            row1,
            text="隐藏到托盘",
            value="tray",
            variable=self._close_behavior_var,
            command=self._persist_settings,
            font=(font_family, 12)
        ).pack(side="left", padx=(14, 0))

        row2 = ctk.CTkFrame(settings, fg_color="transparent")
        row2.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 10))
        
        ctk.CTkLabel(row2, text="开机自启", font=(font_family, 12)).pack(side="left")
        self._autostart_on_radio = ctk.CTkRadioButton(
            row2,
            text="开",
            value="on",
            variable=self._autostart_var,
            command=self._on_autostart_changed,
            font=(font_family, 12)
        )
        self._autostart_on_radio.pack(side="left", padx=(14, 0))
        self._autostart_off_radio = ctk.CTkRadioButton(
            row2,
            text="关",
            value="off",
            variable=self._autostart_var,
            command=self._on_autostart_changed,
            font=(font_family, 12)
        )
        self._autostart_off_radio.pack(side="left", padx=(14, 0))
        
        if not self._is_frozen:
            self._autostart_var.set("off")
            self._autostart_on_radio.configure(state="disabled")
            self._autostart_off_radio.configure(state="disabled")

        row3 = ctk.CTkFrame(settings, fg_color="transparent")
        row3.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 10))

        ctk.CTkLabel(row3, text="窗口", font=(font_family, 12)).pack(side="left")
        self._topmost_button = ctk.CTkButton(
            row3,
            text="置顶窗口",
            command=self._toggle_topmost,
            font=(font_family, 12),
            width=90,
        )
        self._topmost_button.pack(side="left", padx=(14, 0))

        footer = ctk.CTkFrame(outer, fg_color="transparent")
        footer.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 6))
        footer.grid_columnconfigure(0, weight=1)

        footer_content = ctk.CTkFrame(footer, fg_color="transparent")
        footer_content.grid(row=0, column=0)

        ctk.CTkLabel(footer_content, text="by igttttma ", font=(font_family, 11)).pack(side="left")
        link = ctk.CTkLabel(
            footer_content,
            text="项目主页",
            font=(font_family, 11, "underline"),
            text_color=("#1d4ed8", "#60a5fa"),
            cursor="hand2",
        )
        link.pack(side="left")
        link.bind("<Button-1>", lambda _event: self._open_project_homepage())

    def _on_copy_clicked(self) -> None:
        self._set_status("")
        latex = self._text.get("1.0", "end").strip()
        if not latex:
            return
        try:
            mathml = convert(latex)
            copy_text(mathml)
            self._set_status("完成")
        except Exception as e:
            self._set_status(f"失败：{e}")

    def _on_toggle_auto_paste(self) -> None:
        enabled = bool(self._auto_paste_var.get())
        self._auto_paster.set_enabled(enabled)
        self._persist_settings()

    def _toggle_topmost(self) -> None:
        self._set_topmost_enabled(not self._topmost_enabled)

    def _open_project_homepage(self) -> None:
        webbrowser.open_new_tab("https://github.com/igttttma/latex2word")

    def _set_topmost_enabled(self, enabled: bool) -> None:
        enabled = bool(enabled)
        self._topmost_enabled = enabled
        try:
            self._root.attributes("-topmost", enabled)
        except Exception:
            pass
        btn = self._topmost_button
        if btn is not None:
            btn.configure(text="取消置顶" if enabled else "置顶窗口")

    def _set_status(self, text: str) -> None:
        self._status_var.set(text)
        if text == "完成":
            self._root.after(1200, lambda: self._status_var.set(""))

    def _center_window(self) -> None:
        self._root.update_idletasks()

        default_width = 300
        
        req_width = self._root.winfo_reqwidth()
        req_height = self._root.winfo_reqheight()
        
        w = max(req_width, default_width)
        h = req_height
        
        sw = self._root.winfo_screenwidth()
        sh = self._root.winfo_screenheight()
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 2)
        self._root.geometry(f"{w}x{h}+{x}+{y}")
        self._root.minsize(w, h)
        self._root.maxsize(w, h)
        self._centered_once = True

    def _on_close(self) -> None:
        if self._close_behavior_var.get() == "tray":
            self._persist_settings()
            self._root.withdraw()
            return
        self._exit_app()

    def _exit_app(self) -> None:
        self._persist_settings()
        self._auto_paster.stop()
        tray = self._tray
        self._tray = None
        if tray is not None:
            tray.stop()
        self._root.destroy()

    def _show_window(self) -> None:
        if not self._centered_once:
            self._center_window()
        self._root.deiconify()
        self._root.lift()
        try:
            self._root.focus_force()
        except Exception:
            pass

    def _set_auto_paste_enabled(self, enabled: bool) -> None:
        self._auto_paste_var.set(bool(enabled))
        self._on_toggle_auto_paste()

    def _load_settings(self) -> None:
        settings = windows_settings.load_settings(is_frozen=self._is_frozen)
        if settings is None:
            return
        close_behavior = settings.get("close_behavior")
        if close_behavior in {"exit", "tray"}:
            self._close_behavior_var.set(close_behavior)
        autostart = settings.get("autostart")
        if self._is_frozen and autostart in {"on", "off"}:
            self._autostart_var.set(autostart)
        auto_paste = settings.get("auto_paste")
        if isinstance(auto_paste, bool):
            self._auto_paste_var.set(auto_paste)

    def _sync_autostart_state(self) -> None:
        state = windows_settings.get_autostart_state(is_frozen=self._is_frozen)
        if state in {"on", "off"}:
            self._autostart_var.set(state)
        if state == "on":
            if self._close_behavior_var.get() != "tray":
                self._close_behavior_var.set("tray")
                self._persist_settings()
            windows_settings.ensure_autostart_silent(is_frozen=self._is_frozen)

    def _persist_settings(self) -> None:
        windows_settings.persist_settings(
            close_behavior=self._close_behavior_var.get(),
            autostart=self._autostart_var.get(),
            auto_paste=bool(self._auto_paste_var.get()),
            is_frozen=self._is_frozen,
        )

    def _on_autostart_changed(self) -> None:
        self._persist_settings()
        self._apply_autostart()

    def _apply_autostart(self) -> None:
        windows_settings.apply_autostart(autostart=self._autostart_var.get(), is_frozen=self._is_frozen)

    def _create_tray(self) -> None:
        tray = TrayIcon(
            on_show=lambda: self._root.after(0, self._show_window),
            on_toggle_auto_paste=lambda: self._root.after(
                0, lambda: self._set_auto_paste_enabled(not self._auto_paste_var.get())
            ),
            on_exit=lambda: self._root.after(0, self._exit_app),
            get_auto_paste_state=lambda: bool(self._auto_paste_var.get()),
            icon_path=self._icon_ico_path,
        )
        tray.start()
        self._tray = tray
