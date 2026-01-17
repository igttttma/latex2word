from __future__ import annotations

import logging
import re
import sys
from collections.abc import Callable

from src.ui.win_clipboard_watcher import WinClipboardWatcher

logger = logging.getLogger(__name__)


class ClipboardAutoPaster:
    _INITIAL_READ_DELAY_MS = 50

    def __init__(
        self,
        *,
        root,
        get_clipboard_text: Callable[[], str],
        set_clipboard_text: Callable[[str], None],
        convert_latex: Callable[[str], str],
        on_preview: Callable[[str], None],
    ) -> None:
        self._root = root
        self._get_clipboard_text = get_clipboard_text
        self._set_clipboard_text = set_clipboard_text
        self._convert_latex = convert_latex
        self._on_preview = on_preview

        self._read_in_flight = False
        self._enabled = False
        self._win_watcher = (
            WinClipboardWatcher(on_update=lambda: self._root.after(0, self._on_clipboard_update))
            if sys.platform == "win32"
            else None
        )

    def set_enabled(self, enabled: bool) -> None:
        enabled = bool(enabled)
        if enabled and not self._enabled:
            self._enabled = True
            logger.info("auto_paste enabled")
            watcher = self._win_watcher
            if watcher is None:
                self._enabled = False
                logger.warning("auto_paste unavailable on non-Windows platform")
                return
            watcher.start()
            self._on_clipboard_update()
            return
        if not enabled:
            self.stop()

    def stop(self) -> None:
        if self._enabled:
            logger.info("auto_paste disabled")
        self._enabled = False
        watcher = self._win_watcher
        if watcher is not None:
            watcher.stop()

    def _on_clipboard_update(self) -> None:
        if not self._enabled:
            return
        if self._read_in_flight:
            return
        self._read_in_flight = True
        self._root.after(self._INITIAL_READ_DELAY_MS, self._on_clipboard_update_after_delay)

    def _on_clipboard_update_after_delay(self) -> None:
        if not self._enabled:
            self._read_in_flight = False
            return
        try:
            text = self._get_clipboard_text()
        except Exception as e:
            logger.info("clipboard read failed error=%r", e)
            self._read_in_flight = False
            return

        if not text:
            self._read_in_flight = False
            return
        self._read_in_flight = False

        if self._looks_like_latex(text):
            summary = self._summarize(text)
            self._on_preview(summary)
            logger.info("latex detected len=%s summary=%r", len(text), summary)
            try:
                mathml = self._convert_latex(text)
            except Exception as e:
                logger.info("latex convert failed error=%r", e)
                return

            if mathml and "<math" in mathml and mathml != text:
                logger.info("latex converted mathml_len=%s", len(mathml))
                self._try_write_back(mathml=mathml)
        else:
            self._on_preview(self._summarize(text))

    def _try_write_back(self, *, mathml: str) -> None:
        if not self._enabled:
            return
        try:
            self._set_clipboard_text(mathml)
        except Exception as e:
            logger.info("clipboard write failed error=%r", e)
            return

        logger.info("write_back success")

    def _looks_like_latex(self, text: str) -> bool:
        s = text.strip()
        if not s:
            return False
        if len(s) > 50000:
            return False
        if "\\mathrm" in s:
            return False
        if s.startswith("<math"):
            return False
        if 'xmlns="http://www.w3.org/1998/Math/MathML"' in s:
            return False
        if re.search(r"\bhttps?://", s):
            return False

        if re.search(r"\b[A-Za-z]:\\", s):
            return False

        if "$" in s:
            if s.startswith("$$") or s.endswith("$$"):
                if not (s.startswith("$$") and s.endswith("$$")):
                    return False
                inner = s[2:-2]
                if not inner.strip():
                    return False
                if re.search(r"(?<!\\)\$", inner):
                    return False
                return True

            if s.startswith("$") and s.endswith("$"):
                inner = s[1:-1]
                if not inner.strip():
                    return False
                if re.search(r"(?<!\\)\$", inner):
                    return False
                return True
        has_tex_command = bool(re.search(r"\\[a-zA-Z]{2,}\b", s))
        if has_tex_command:
            return True
        has_structural_chars = any(ch in s for ch in ("{", "}", "\\\\", "&"))
        if has_structural_chars:
            return True
        has_script = bool(re.search(r"[_^](?:\{[^}]+\}|\w+)", s))
        if has_script:
            if re.fullmatch(r"[A-Za-z0-9_]+", s) and s.count("_") >= 2:
                return False
            if len(s) <= 60:
                return True
            space_runs = len(re.findall(r"\s+", s))
            has_sentence_punct = bool(re.search(r"[.?!;:。？！；：]", s))
            if space_runs >= 12:
                return False
            if has_sentence_punct and space_runs >= 4:
                return False
            if len(s) >= 120 and space_runs >= 4:
                return False
            return True
        return False

    def _summarize(self, text: str) -> str:
        s = re.sub(r"\s+", " ", text.strip())
        if len(s) <= 140:
            return s
        return s[:140] + "…"
