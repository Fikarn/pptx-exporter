"""ErrorBanner — inline error display with red left-border accent."""

import customtkinter as ctk

from ..tokens import COLORS, FONTS, RADIUS, SP


class ErrorBanner(ctk.CTkFrame):
    """Inline error banner shown below the action area on failure."""

    def __init__(self, parent, on_dismiss):
        super().__init__(parent, fg_color="transparent", corner_radius=0)
        self.grid_columnconfigure(0, weight=0)  # accent bar
        self.grid_columnconfigure(1, weight=1)  # message
        self.grid_columnconfigure(2, weight=0)  # dismiss
        self._on_dismiss = on_dismiss

        # Red left accent bar
        self._accent = ctk.CTkFrame(
            self, width=4, fg_color=COLORS["error"], corner_radius=0,
        )
        self._accent.grid(row=0, column=0, sticky="ns")

        # Message area
        msg_frame = ctk.CTkFrame(
            self, fg_color=COLORS["error_bg"],
            corner_radius=0,
        )
        msg_frame.grid(row=0, column=1, sticky="nsew")
        msg_frame.grid_columnconfigure(0, weight=1)

        self._msg_label = ctk.CTkLabel(
            msg_frame,
            text="",
            font=FONTS["caption"],
            text_color=COLORS["error_text"],
            anchor="w",
            wraplength=420,
            justify="left",
        )
        self._msg_label.grid(
            row=0, column=0, sticky="ew",
            padx=SP["sm"], pady=SP["sm"],
        )

        ctk.CTkButton(
            msg_frame,
            text="\u2715",
            command=self._on_dismiss,
            width=24,
            height=24,
            font=FONTS["caption"],
            fg_color="transparent",
            text_color=COLORS["error_text"],
            hover_color=COLORS["error_bg"],
            corner_radius=RADIUS["sm"],
        ).grid(row=0, column=1, padx=(0, SP["sm"]), pady=SP["sm"])

    def show(self, message: str) -> None:
        self._msg_label.configure(text=message)
