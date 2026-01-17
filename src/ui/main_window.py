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
        self._topmost_button: ctk.CTkSwitch | None = None
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
        font_family = "Microsoft YaHei UI"
        
        self._root.grid_columnconfigure(0, weight=1)
        self._root.grid_rowconfigure(0, weight=1)

        outer = ctk.CTkFrame(self._root, fg_color="transparent")
        outer.grid(row=0, column=0, sticky="nsew", padx=12, pady=12)
        outer.grid_columnconfigure(0, weight=1)

        card_color = ("white", "#2b2b2b")
        card_border_color = ("#e5e5e5", "#3d3d3d")
        card_border_width = 1
        card_corner_radius = 10
        section_padding = (0, 8)

        manual = ctk.CTkFrame(
            outer, 
            fg_color=card_color, 
            corner_radius=card_corner_radius,
            border_width=card_border_width,
            border_color=card_border_color
        )
        manual.grid(row=0, column=0, sticky="ew", padx=0, pady=section_padding)
        manual.grid_columnconfigure(0, weight=1)
        
        manual_label = ctk.CTkLabel(manual, text="手动输入", font=(font_family, 14, "bold"))
        manual_label.grid(row=0, column=0, sticky="w", padx=12, pady=(10, 3))

        self._text = ctk.CTkTextbox(
            manual,
            width=300,
            height=56,
            font=(font_family, 12),
            wrap="word",
            undo=True,
            fg_color=("gray98", "gray20"),
            border_width=1,
            border_color=("gray85", "gray30")
        )
        self._text.grid(row=1, column=0, sticky="ew", padx=12, pady=6)

        manual_bottom = ctk.CTkFrame(manual, fg_color="transparent")
        manual_bottom.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 10))
        manual_bottom.grid_columnconfigure(0, weight=1)

        self._status_label = ctk.CTkLabel(manual_bottom, textvariable=self._status_var, font=(font_family, 12), text_color=("gray50", "gray70"))
        self._status_label.grid(row=0, column=0, sticky="w")

        copy_btn = ctk.CTkButton(
            manual_bottom, 
            text="复制结果", 
            command=self._on_copy_clicked,
            font=(font_family, 12),
            height=30,
            corner_radius=8,
            width=90,
        )
        copy_btn.grid(row=0, column=1, sticky="e")

        auto = ctk.CTkFrame(
            outer, 
            fg_color=card_color, 
            corner_radius=card_corner_radius,
            border_width=card_border_width,
            border_color=card_border_color
        )
        auto.grid(row=1, column=0, sticky="ew", padx=0, pady=section_padding)
        auto.grid_columnconfigure(0, weight=1)
        
        auto_header = ctk.CTkFrame(auto, fg_color="transparent")
        auto_header.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 3))
        auto_header.grid_columnconfigure(0, weight=1)

        auto_label = ctk.CTkLabel(auto_header, text="自动识别", font=(font_family, 14, "bold"))
        auto_label.grid(row=0, column=0, sticky="w")

        self._auto_paste_check = ctk.CTkSwitch(
            auto_header,
            text="开启",
            variable=self._auto_paste_var,
            command=self._on_toggle_auto_paste,
            font=(font_family, 12),
            width=50,
            button_color="#3b8ed0",
            progress_color="#3b8ed0"
        )
        self._auto_paste_check.grid(row=0, column=1, sticky="e")

        desc = ctk.CTkLabel(
            auto,
            text="监听剪贴板，自动识别 LaTeX 公式并转换。",
            wraplength=320,
            justify="left",
            font=(font_family, 12),
            text_color=("gray40", "gray60")
        )
        desc.grid(row=1, column=0, sticky="w", padx=12, pady=(0, 6))

        preview = ctk.CTkEntry(
            auto, 
            textvariable=self._auto_paste_preview_var, 
            state="readonly", 
            font=(font_family, 12),
            height=30,
            border_color=("gray85", "gray30"),
            fg_color=("gray98", "gray20")
        )
        preview.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 10))

        settings = ctk.CTkFrame(
            outer, 
            fg_color=card_color, 
            corner_radius=card_corner_radius,
            border_width=card_border_width,
            border_color=card_border_color
        )
        settings.grid(row=2, column=0, sticky="ew", padx=0, pady=section_padding)
        settings.grid_columnconfigure(0, weight=1)
        
        settings_label = ctk.CTkLabel(settings, text="设置", font=(font_family, 14, "bold"))
        settings_label.grid(row=0, column=0, sticky="w", padx=12, pady=(10, 6))

        row1 = ctk.CTkFrame(settings, fg_color="transparent")
        row1.grid(row=1, column=0, sticky="ew", padx=12, pady=3)
        
        ctk.CTkLabel(row1, text="关闭窗口时", font=(font_family, 12)).pack(side="left")
        
        radio_style = {"font": (font_family, 12), "border_width_checked": 4, "border_width_unchecked": 2}
        
        ctk.CTkRadioButton(
            row1, text="退出程序", value="exit", variable=self._close_behavior_var,
            command=self._persist_settings, **radio_style
        ).pack(side="right", padx=(10, 0))
        
        ctk.CTkRadioButton(
            row1, text="最小化到托盘", value="tray", variable=self._close_behavior_var,
            command=self._persist_settings, **radio_style
        ).pack(side="right", padx=(10, 0))

        row2 = ctk.CTkFrame(settings, fg_color="transparent")
        row2.grid(row=2, column=0, sticky="ew", padx=12, pady=3)

        ctk.CTkLabel(row2, text="开机自启", font=(font_family, 12)).pack(side="left")

        if self._is_frozen:
            ctk.CTkRadioButton(
                row2,
                text="关",
                value="off",
                variable=self._autostart_var,
                command=self._on_autostart_changed,
                **radio_style,
            ).pack(side="right", padx=(10, 0))

            ctk.CTkRadioButton(
                row2,
                text="开",
                value="on",
                variable=self._autostart_var,
                command=self._on_autostart_changed,
                **radio_style,
            ).pack(side="right", padx=(10, 0))
        else:
            ctk.CTkLabel(
                row2,
                text="开机自启仅对exe生效，当前为py启动",
                font=(font_family, 12),
                text_color=("gray50", "gray70"),
            ).pack(side="right")

        row3 = ctk.CTkFrame(settings, fg_color="transparent")
        row3.grid(row=3, column=0, sticky="ew", padx=12, pady=(3, 10))

        ctk.CTkLabel(row3, text="窗口置顶", font=(font_family, 12)).pack(side="left")
        self._topmost_button = ctk.CTkSwitch(
            row3,
            text="",
            command=self._toggle_topmost,
            font=(font_family, 12),
            width=50,
            onvalue=True,
            offvalue=False
        )
        self._topmost_button.pack(side="right", padx=(0, 0))
        
        footer = ctk.CTkFrame(outer, fg_color="transparent")
        footer.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 2))
        footer.grid_columnconfigure(0, weight=1)

        footer_content = ctk.CTkFrame(footer, fg_color="transparent")
        footer_content.pack(expand=True)

        ctk.CTkLabel(footer_content, text="By igttttma", font=(font_family, 11), text_color="gray60").pack(side="left", padx=(0, 5))
        link = ctk.CTkLabel(
            footer_content,
            text="GitHub",
            font=(font_family, 11),
            text_color=("#3b8ed0", "#60a5fa"),
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
        btn = self._topmost_button
        if btn is None:
            return
        self._set_topmost_enabled(bool(btn.get()))

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
            current = bool(btn.get())
            if current != enabled:
                if enabled:
                    btn.select()
                else:
                    btn.deselect()

    def _set_status(self, text: str) -> None:
        self._status_var.set(text)
        if text == "完成":
            self._root.after(1200, lambda: self._status_var.set(""))

    def _center_window(self) -> None:
        self._root.update_idletasks()

        default_width = 380
        default_height = 495
        
        sw = self._root.winfo_screenwidth()
        sh = self._root.winfo_screenheight()

        w = default_width
        h = default_height
        w = min(w, max(320, sw - 80))
        h = min(h, max(240, sh - 120))
        
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
