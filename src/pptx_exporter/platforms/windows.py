"""Windows backend: drive Microsoft PowerPoint via win32com COM automation.

Per-slide workflow (mirrors manual user actions):
1. Add a transparent, borderless full-slide rectangle via COM.
2. Select all objects on the slide (ctrl+a equivalent via SelectAll).
3. Export the selection as a transparent PNG via ExportAsFixedFormat / SaveAs.
4. Remove the bounding rectangle.
"""

import logging
import threading
from pathlib import Path
from typing import Callable, Optional

from ..utils import slide_output_name

logger = logging.getLogger(__name__)

# PowerPoint constants (from Microsoft Office type library)
_PP_SAVE_AS_PNG = 32          # PpSaveAsFileType.ppSaveAsPNG
_MSO_FALSE = 0
_MSO_TRUE = -1
_PP_MEDIA_TYPE_PICTURE = 0   # ppMediaTypePicture (for export)

# AutoShape rectangle
_MSO_SHAPE_RECTANGLE = 1     # msoShapeRectangle


def export_slides(
    pptx_path: Path,
    output_dir: Path,
    progress_callback: Optional[Callable[[int, int], None]] = None,
    cancel_event: Optional[threading.Event] = None,
) -> None:
    """Export every slide of *pptx_path* as a transparent PNG into *output_dir*.

    Uses win32com to drive Microsoft PowerPoint for Windows.

    Args:
        pptx_path: Resolved path to the .pptx file.
        output_dir: Resolved path to the output directory.
        progress_callback: Optional callable(slide_index, total_slides).
    """
    try:
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required for the Windows backend. "
            "Install it with: pip install pywin32"
        ) from exc

    pptx_str = str(pptx_path)
    logger.info("Opening presentation via COM: %s", pptx_str)

    app = win32.Dispatch("PowerPoint.Application")
    app.Visible = _MSO_FALSE  # run headlessly

    try:
        prs = app.Presentations.Open(
            pptx_str,
            ReadOnly=_MSO_FALSE,
            Untitled=_MSO_FALSE,
            WithWindow=_MSO_FALSE,
        )

        total = prs.Slides.Count
        slide_w_pts = prs.PageSetup.SlideWidth
        slide_h_pts = prs.PageSetup.SlideHeight
        logger.info(
            "Slide count: %d, dimensions: %.1f x %.1f pts",
            total,
            slide_w_pts,
            slide_h_pts,
        )

        for idx in range(total):
            if cancel_event and cancel_event.is_set():
                logger.info("Export cancelled before slide %d", idx + 1)
                raise InterruptedError("Export cancelled by user.")

            slide_num = idx + 1
            if progress_callback:
                progress_callback(idx, total)

            slide = prs.Slides(slide_num)
            logger.debug("Processing slide %d/%d", slide_num, total)

            out_name = slide_output_name(idx, total)
            out_path = str(output_dir / out_name)

            # Step 1 — add invisible bounding rectangle
            bounding = slide.Shapes.AddShape(
                _MSO_SHAPE_RECTANGLE,
                0,          # Left
                0,          # Top
                slide_w_pts,
                slide_h_pts,
            )
            bounding.Fill.Visible = _MSO_FALSE
            bounding.Line.Visible = _MSO_FALSE
            bounding_id = bounding.Id
            logger.debug("Added bounding rect, shape id: %d", bounding_id)

            # Step 2+3 — export all shapes as PNG via ShapeRange.Export
            try:
                shape_range = slide.Shapes.Range()
                # Export the full range as PNG
                # ppShapeFormatPNG = 2
                shape_range.Export(out_path, 2)
                logger.info("Exported slide %d → %s", slide_num, out_path)
            except Exception as exc:
                logger.warning(
                    "ShapeRange.Export failed for slide %d: %s — falling back to slide export",
                    slide_num,
                    exc,
                )
                # Fallback: export the entire slide (not transparent, but better than crashing)
                slide.Export(out_path, "PNG")

            # Step 4 — remove bounding rectangle
            try:
                for i in range(1, slide.Shapes.Count + 1):
                    if slide.Shapes(i).Id == bounding_id:
                        slide.Shapes(i).Delete()
                        break
                logger.debug("Removed bounding rect from slide %d", slide_num)
            except Exception as exc:
                logger.warning(
                    "Could not remove bounding rect from slide %d: %s", slide_num, exc
                )

    finally:
        try:
            prs.Close()
            logger.info("Closed presentation")
        except Exception as exc:
            logger.warning("Could not close presentation: %s", exc)
        try:
            app.Quit()
            logger.info("Quit PowerPoint")
        except Exception as exc:
            logger.warning("Could not quit PowerPoint: %s", exc)

    if progress_callback:
        progress_callback(total, total)
