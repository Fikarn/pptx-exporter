"""DropZone — large drag-and-drop target with browse button (empty file state)."""

import customtkinter as ctk

from ..tokens import COLORS, FONTS, SP


class DropZone(ctk.CTkFrame):
    """Empty-state file panel: large drop target with icon and browse button."""

    def __init__(self, parent, on_browse, dnd_enabled: bool = False):
        super().__init__(parent, fg_color="transparent")
        self.grid_columnconfigure(0, weight=1)
        self._on_browse = on_browse

        # Drop icon
        ctk.CTkLabel(
            self,
            text="\u2B07",
            font=("system-ui", 28),
            text_color=COLORS["text_tertiary"],
        ).grid(row=0, column=0, pady=(SP["lg"], SP["xs"]))

        # Primary text
        ctk.CTkLabel(
            self,
            text="Drop .pptx files here",
            font=FONTS["body_bold"],
            text_color=COLORS["text_primary"],
        ).grid(row=1, column=0, pady=(0, SP["xs"]))

        # Secondary text
        subline = (
            "or click Browse to select files"
            if dnd_enabled
            else "Click Browse to select .pptx files"
        )
        ctk.CTkLabel(
            self,
            text=subline,
            font=FONTS["caption"],
            text_color=COLORS["text_secondary"],
        ).grid(row=2, column=0, pady=(0, SP["md"]))

        # Browse button
        ctk.CTkButton(
            self,
            text="Browse Files\u2026",
            command=self._on_browse,
            width=130,
            height=34,
            font=FONTS["body"],
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            corner_radius=8,
        ).grid(row=3, column=0, pady=(0, SP["lg"]))
