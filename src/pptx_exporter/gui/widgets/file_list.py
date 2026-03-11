"""FileList — scrollable list of loaded files with remove buttons."""

from pathlib import Path

import customtkinter as ctk

from ..tokens import COLORS, FONTS, SP


class FileList(ctk.CTkFrame):
    """Loaded-state file panel: scrollable file list with add/remove controls."""

    # Switch to scrollable frame when file count exceeds this
    _SCROLL_THRESHOLD = 4

    def __init__(self, parent, on_browse, on_clear_all, on_remove_file):
        super().__init__(parent, fg_color="transparent")
        self.grid_columnconfigure(0, weight=1)
        self._on_browse = on_browse
        self._on_clear_all = on_clear_all
        self._on_remove_file = on_remove_file
        self._file_rows: dict[str, ctk.CTkFrame] = {}

        # List container (replaced with scrollable when needed)
        self._list_container = ctk.CTkFrame(self, fg_color="transparent")
        self._list_container.grid(
            row=0, column=0, sticky="ew", padx=SP["md"], pady=(SP["sm"], 0),
        )
        self._list_container.grid_columnconfigure(0, weight=1)
        self._scrollable = None

        # Divider
        ctk.CTkFrame(
            self, height=1, fg_color=COLORS["border"], corner_radius=0,
        ).grid(row=1, column=0, sticky="ew", padx=SP["md"], pady=(SP["sm"], 0))

        # Bottom bar
        btn_bar = ctk.CTkFrame(self, fg_color="transparent")
        btn_bar.grid(
            row=2, column=0, sticky="ew",
            padx=SP["md"], pady=(SP["sm"], SP["sm"]),
        )
        btn_bar.grid_columnconfigure(0, weight=1)

        ctk.CTkButton(
            btn_bar,
            text="+ Add Files",
            command=self._on_browse,
            width=100,
            height=28,
            font=FONTS["caption"],
            fg_color="transparent",
            text_color=COLORS["accent"],
            hover_color=COLORS["surface_hover"],
            border_width=1,
            border_color=COLORS["accent"],
            corner_radius=6,
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkButton(
            btn_bar,
            text="Clear All",
            command=self._on_clear_all,
            width=80,
            height=28,
            font=FONTS["caption"],
            fg_color="transparent",
            text_color=COLORS["text_secondary"],
            hover_color=COLORS["surface_hover"],
            border_width=1,
            border_color=COLORS["border"],
            corner_radius=6,
        ).grid(row=0, column=1, sticky="e")

    def set_files(self, file_infos: list[tuple[str, int]]) -> None:
        """Rebuild the file list from a list of (path, slide_count) tuples."""
        # Clear old rows
        for widget in self._file_rows.values():
            widget.destroy()
        self._file_rows.clear()

        # Decide whether to use scrollable container
        use_scroll = len(file_infos) > self._SCROLL_THRESHOLD

        if self._scrollable is not None:
            self._scrollable.destroy()
            self._scrollable = None

        if use_scroll:
            self._list_container.grid_remove()
            self._scrollable = ctk.CTkScrollableFrame(
                self,
                fg_color="transparent",
                height=140,
            )
            self._scrollable.grid(
                row=0, column=0, sticky="ew",
                padx=SP["md"], pady=(SP["sm"], 0),
            )
            self._scrollable.grid_columnconfigure(0, weight=1)
            target = self._scrollable
        else:
            self._list_container.grid()
            target = self._list_container

        for i, (path, slide_count) in enumerate(file_infos):
            self._add_row(target, i, path, slide_count)

    def _add_row(
        self, parent, index: int, path: str, slide_count: int,
    ) -> None:
        row = ctk.CTkFrame(parent, fg_color="transparent", height=30)
        row.grid(row=index, column=0, sticky="ew", pady=(0, 2))
        row.grid_columnconfigure(1, weight=1)

        # File icon
        ctk.CTkLabel(
            row, text="\U0001F4C4", font=("system-ui", 13),
            width=20,
        ).grid(row=0, column=0, padx=(0, SP["xs"]))

        # Filename
        ctk.CTkLabel(
            row,
            text=Path(path).name,
            font=FONTS["body"],
            text_color=COLORS["text_primary"],
            anchor="w",
        ).grid(row=0, column=1, sticky="w")

        # Slide count
        if slide_count > 0:
            text = f"{slide_count} slide{'s' if slide_count != 1 else ''}"
            ctk.CTkLabel(
                row,
                text=text,
                font=FONTS["caption"],
                text_color=COLORS["text_tertiary"],
                anchor="e",
            ).grid(row=0, column=2, padx=(SP["sm"], 0))

        # Remove button
        p = path
        ctk.CTkButton(
            row,
            text="\u2715",
            command=lambda: self._on_remove_file(p),
            width=24,
            height=24,
            font=FONTS["caption"],
            fg_color="transparent",
            text_color=COLORS["text_tertiary"],
            hover_color=COLORS["surface_hover"],
            corner_radius=4,
        ).grid(row=0, column=3, padx=(SP["xs"], 0))

        self._file_rows[path] = row
