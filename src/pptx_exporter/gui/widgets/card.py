"""Card — rounded surface container with optional border."""

import customtkinter as ctk

from ..tokens import COLORS, SP


class Card(ctk.CTkFrame):
    """A rounded-corner card container used as the base surface for sections."""

    def __init__(self, parent, **kwargs):
        defaults = dict(
            fg_color=COLORS["surface"],
            border_color=COLORS["border"],
            border_width=1,
            corner_radius=12,
        )
        defaults.update(kwargs)
        super().__init__(parent, **defaults)
        self.grid_columnconfigure(0, weight=1)
        self._pad = SP["md"]

    def content_pad(self) -> dict:
        """Return standard padx/pady for content inside the card."""
        return {"padx": self._pad, "pady": self._pad}
