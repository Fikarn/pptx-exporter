"""Core export logic — OS-agnostic interface.

Detects the active backend at construction time and delegates per-slide
export to the appropriate platform module.
"""

import logging
import threading
from pathlib import Path
from typing import Callable, Optional

from .utils import (
    backend_description,
    detect_backend,
    validate_output_dir,
    validate_pptx,
)

logger = logging.getLogger(__name__)


class Exporter:
    """High-level export controller.

    Attributes:
        backend: One of ``'macos'``, ``'windows'``, or ``'not_found'``.
        backend_label: Human-readable description of the active backend.
    """

    def __init__(self) -> None:
        self.backend: str = detect_backend()
        self.backend_label: str = backend_description(self.backend)
        logger.info("Active backend: %s", self.backend_label)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def export(
        self,
        pptx_path: str,
        output_dir: str,
        progress_callback: Optional[Callable[[int, int], None]] = None,
        cancel_event: Optional[threading.Event] = None,
        ppi: int = 300,
    ) -> None:
        """Validate inputs and run the export.

        Args:
            pptx_path: Path to the source .pptx file.
            output_dir: Path to the destination directory.
            progress_callback: Optional callable(current_slide_index, total_slides)
                called before each slide is processed, and once more at
                completion with current_slide_index == total_slides.
            cancel_event: Optional threading.Event; when set the export is
                aborted cleanly between slides and InterruptedError is raised.
            ppi: Output resolution in pixels per inch (default 300).

        Raises:
            ValueError: If *pptx_path* or *output_dir* are invalid.
            RuntimeError: If the underlying backend fails.
            InterruptedError: If *cancel_event* is set during the export.
        """
        pptx = validate_pptx(pptx_path)
        out = validate_output_dir(output_dir)

        logger.info(
            "Starting export: pptx=%s  output=%s  backend=%s  ppi=%d",
            pptx,
            out,
            self.backend,
            ppi,
        )

        if self.backend == "not_found":
            raise RuntimeError(
                "Microsoft PowerPoint is required but was not found. "
                "Please install PowerPoint and try again."
            )

        if self.backend == "macos":
            self._export_macos(pptx, out, progress_callback, cancel_event, ppi)
        elif self.backend == "windows":
            self._export_windows(pptx, out, progress_callback, cancel_event, ppi)

        logger.info("Export complete → %s", out)

    # ------------------------------------------------------------------
    # Private dispatch helpers
    # ------------------------------------------------------------------

    def _export_macos(
        self,
        pptx: Path,
        out: Path,
        progress_callback: Optional[Callable[[int, int], None]],
        cancel_event: Optional[threading.Event],
        ppi: int,
    ) -> None:
        from .platforms.macos import export_slides

        export_slides(pptx, out, progress_callback=progress_callback,
                      cancel_event=cancel_event, ppi=ppi)

    def _export_windows(
        self,
        pptx: Path,
        out: Path,
        progress_callback: Optional[Callable[[int, int], None]],
        cancel_event: Optional[threading.Event],
        ppi: int,
    ) -> None:
        from .platforms.windows import export_slides

        export_slides(pptx, out, progress_callback=progress_callback,
                      cancel_event=cancel_event, ppi=ppi)
