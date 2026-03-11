"""StatusPill — compact backend status indicator."""

import customtkinter as ctk

from ..tokens import COLORS, FONTS


class StatusPill(ctk.CTkLabel):
    """Small pill showing PowerPoint connection status."""

    def __init__(self, parent, **kwargs):
        super().__init__(
            parent,
            text="",
            font=FONTS["caption"],
            corner_radius=10,
            padx=10,
            pady=3,
            **kwargs,
        )

    def set_ready(self) -> None:
        self.configure(
            text="\u25cf PowerPoint ready",
            fg_color=COLORS["success_bg"],
            text_color=COLORS["success_text"],
        )

    def set_error(self) -> None:
        self.configure(
            text="\u25cf PowerPoint not found",
            fg_color=COLORS["error_bg"],
            text_color=COLORS["error_text"],
        )
