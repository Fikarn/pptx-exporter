"""FilePanel — orchestrates DropZone (empty) and FileList (loaded) states."""

from ..tokens import COLORS
from .card import Card
from .drop_zone import DropZone
from .file_list import FileList


class FilePanel(Card):
    """Top-level file selection card — switches between empty and loaded states."""

    def __init__(self, parent, on_browse, on_clear_all, on_remove_file,
                 dnd_enabled: bool = False):
        super().__init__(parent)
        self._on_browse = on_browse
        self._dnd_enabled = dnd_enabled

        self._drop_zone = DropZone(self, on_browse, dnd_enabled)
        self._file_list = FileList(self, on_browse, on_clear_all, on_remove_file)

        self._drop_zone.grid(row=0, column=0, sticky="ew")
        self._file_list.grid(row=0, column=0, sticky="ew")
        self._file_list.grid_remove()

    def show_empty(self) -> None:
        """Switch to the empty/drop-zone state."""
        self._file_list.grid_remove()
        self._drop_zone.grid()

    def set_files(self, file_infos: list[tuple[str, int]]) -> None:
        """Switch to loaded state and display the given files."""
        if not file_infos:
            self.show_empty()
            return
        self._file_list.set_files(file_infos)
        self._drop_zone.grid_remove()
        self._file_list.grid()

    def set_drop_highlight(self, active: bool) -> None:
        """Toggle accent border for drag-over feedback."""
        self.configure(
            border_color=COLORS["accent"] if active else COLORS["border"],
            border_width=2 if active else 1,
        )
