"""StatusPill — compact backend status indicator with dot + text."""

import customtkinter as ctk

from ..tokens import COLORS, FONTS


class StatusPill(ctk.CTkFrame):
    """Small indicator showing PowerPoint connection status."""

    def __init__(self, parent, **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)

        self._dot = ctk.CTkLabel(
            self, text="\u25CF", font=("system-ui", 8), width=12,
        )
        self._dot.grid(row=0, column=0, padx=(0, 4))

        self._label = ctk.CTkLabel(
            self, text="", font=FONTS["caption"],
            text_color=COLORS["text_secondary"],
        )
        self._label.grid(row=0, column=1)

    def set_ready(self) -> None:
        self._dot.configure(text_color=COLORS["accent"])
        self._label.configure(text="PowerPoint connected")

    def set_error(self) -> None:
        self._dot.configure(text_color=COLORS["error"])
        self._label.configure(
            text="PowerPoint not found",
            text_color=COLORS["error_text"],
        )
