"""DropZone — large drag-and-drop target with browse button (empty file state)."""

import customtkinter as ctk

from ..tokens import COLORS, FONTS, RADIUS, SP


class DropZone(ctk.CTkFrame):
    """Empty-state file panel: large drop target with icon and browse button.

    Centers its content vertically when the parent card is taller than needed.
    """

    def __init__(self, parent, on_browse, dnd_enabled: bool = False):
        super().__init__(parent, fg_color="transparent")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)  # center inner frame vertically
        self._on_browse = on_browse

        # Inner frame — bordered area that stays centered
        inner = ctk.CTkFrame(
            self,
            fg_color="transparent",
            border_color=COLORS["border"],
            border_width=1,
            corner_radius=RADIUS["md"],
        )
        inner.grid(
            row=0, column=0, sticky="ew",
            padx=SP["sm"], pady=SP["sm"],
        )
        inner.grid_columnconfigure(0, weight=1)

        # Icon — subtle arrow
        ctk.CTkLabel(
            inner,
            text="\u2193",
            font=(FONTS["display"][0], 24),
            text_color=COLORS["text_tertiary"],
        ).grid(row=0, column=0, pady=(SP["2xl"], SP["sm"]))

        # Primary text — uppercase, editorial
        ctk.CTkLabel(
            inner,
            text="DROP FILES HERE",
            font=FONTS["heading"],
            text_color=COLORS["text_secondary"],
        ).grid(row=1, column=0, pady=(0, SP["xs"]))

        # Secondary text
        subline = (
            "Drag .pptx files or click browse"
            if dnd_enabled
            else "Click browse to select .pptx files"
        )
        ctk.CTkLabel(
            inner,
            text=subline,
            font=FONTS["body_italic"],
            text_color=COLORS["text_tertiary"],
        ).grid(row=2, column=0, pady=(0, SP["md"]))

        # Browse button — ghost/outline style
        ctk.CTkButton(
            inner,
            text="Browse\u2026",
            command=self._on_browse,
            width=110,
            height=32,
            font=FONTS["body"],
            fg_color="transparent",
            text_color=COLORS["accent"],
            hover_color=COLORS["accent_muted"],
            border_width=1,
            border_color=COLORS["accent"],
            corner_radius=RADIUS["sm"],
        ).grid(row=3, column=0, pady=(0, SP["xl"]))
