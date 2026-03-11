"""Windows backend: drive Microsoft PowerPoint via win32com COM automation.

Per-slide workflow:
1. Add a transparent, borderless full-slide rectangle via COM.
2. Copy all shapes to the clipboard via ShapeRange.Copy().
3. Read the clipboard image (PNG or DIB with alpha), resize to target PPI
   via Pillow, and save as a transparent PNG.
4. Remove the bounding rectangle.

Falls back to ShapeRange.Export if the clipboard approach fails.
"""

import io
import logging
import struct
import threading
import time
from pathlib import Path
from typing import Callable, Optional, Sequence

from ..utils import slide_output_name

logger = logging.getLogger(__name__)

# PowerPoint / Office constants
_MSO_FALSE = 0
_MSO_TRUE = -1
_MSO_SHAPE_RECTANGLE = 1  # msoShapeRectangle


# ---------------------------------------------------------------------------
# Clipboard helpers
# ---------------------------------------------------------------------------

def _dib_to_rgba(dib_data: bytes):
    """Convert CF_DIB clipboard data to a Pillow RGBA Image.

    CF_DIB is a BITMAPINFOHEADER followed by pixel data.  32 bpp bitmaps
    use BGRA byte order and may contain a valid alpha channel.
    """
    from PIL import Image

    header_size = struct.unpack_from("<I", dib_data, 0)[0]
    width = struct.unpack_from("<i", dib_data, 4)[0]
    height = struct.unpack_from("<i", dib_data, 8)[0]
    bpp = struct.unpack_from("<H", dib_data, 14)[0]

    bottom_up = height > 0
    height = abs(height)

    if bpp == 32:
        pixel_data = dib_data[header_size:]
        # DIB rows are 4-byte aligned (already true for 32 bpp).
        stride = width * 4
        rows = [pixel_data[y * stride:(y + 1) * stride] for y in range(height)]
        if bottom_up:
            rows.reverse()
        img = Image.frombytes("RGBA", (width, height), b"".join(rows), "raw", "BGRA")
        return img

    # Non-32 bpp — reconstruct a full BMP file so Pillow can open it.
    file_header = struct.pack(
        "<2sIHHI", b"BM", 14 + len(dib_data), 0, 0, 14 + header_size,
    )
    return Image.open(io.BytesIO(file_header + dib_data)).convert("RGBA")


def _read_clipboard_image():
    """Read an image from the Windows clipboard, preserving transparency.

    Prefers the ``PNG`` registered clipboard format (natively supports alpha).
    Falls back to ``CF_DIB`` which may contain 32 bpp BGRA data.

    Returns a Pillow RGBA Image, or *None* if no image is available.
    """
    import win32clipboard

    cf_png = win32clipboard.RegisterClipboardFormat("PNG")

    win32clipboard.OpenClipboard()
    try:
        if win32clipboard.IsClipboardFormatAvailable(cf_png):
            from PIL import Image
            data = win32clipboard.GetClipboardData(cf_png)
            return Image.open(io.BytesIO(data)).convert("RGBA")

        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_DIB):
            data = win32clipboard.GetClipboardData(win32clipboard.CF_DIB)
            return _dib_to_rgba(data)

        return None
    finally:
        win32clipboard.CloseClipboard()


def _save_clipboard_as_png(out_path: str, target_w: int, target_h: int) -> bool:
    """Read clipboard image, resize to *target_w* x *target_h*, save as PNG.

    Returns True on success, False if no image data was found.
    """
    from PIL import Image

    img = _read_clipboard_image()
    if img is None:
        return False

    if img.size != (target_w, target_h):
        img = img.resize((target_w, target_h), Image.LANCZOS)

    img.save(out_path, "PNG")
    return True


# ---------------------------------------------------------------------------
# Main export
# ---------------------------------------------------------------------------

def export_slides(
    pptx_path: Path,
    output_dir: Path,
    progress_callback: Optional[Callable[[int, int], None]] = None,
    cancel_event: Optional[threading.Event] = None,
    ppi: int = 300,
    slide_indices: Optional[Sequence[int]] = None,
) -> None:
    """Export every slide of *pptx_path* as a transparent PNG into *output_dir*.

    Uses win32com to drive Microsoft PowerPoint for Windows.
    Primary method: copy shapes to clipboard → read with Pillow → save PNG.
    Fallback: ShapeRange.Export (may not preserve transparency).

    Args:
        pptx_path: Resolved path to the .pptx file.
        output_dir: Resolved path to the output directory.
        progress_callback: Optional callable(slide_index, total_slides).
        cancel_event: Optional threading.Event; when set the export is
            aborted cleanly between slides and InterruptedError is raised.
        ppi: Output resolution in pixels per inch (default 300).
    """
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required for the Windows backend. "
            "Install it with: pip install pywin32"
        ) from exc

    # COM requires CoInitialize on non-main threads (the GUI runs exports
    # in a background thread).  CoInitialize is a no-op if already called.
    pythoncom.CoInitialize()

    pptx_str = str(pptx_path)
    logger.info("Opening presentation via COM: %s", pptx_str)

    app = win32.Dispatch("PowerPoint.Application")
    app.Visible = _MSO_TRUE
    app.WindowState = 2  # ppWindowMinimized — clipboard requires a window

    try:
        prs = app.Presentations.Open(
            pptx_str,
            ReadOnly=_MSO_FALSE,
            Untitled=_MSO_FALSE,
            WithWindow=_MSO_TRUE,
        )

        total = prs.Slides.Count
        slide_w_pts = prs.PageSetup.SlideWidth
        slide_h_pts = prs.PageSetup.SlideHeight
        # Target pixel dimensions at the requested PPI (points are 1/72 inch).
        target_w = int(round(slide_w_pts / 72 * ppi))
        target_h = int(round(slide_h_pts / 72 * ppi))
        logger.info(
            "Slide count: %d, dimensions: %.1f x %.1f pts, target: %d x %d px (%d PPI)",
            total,
            slide_w_pts,
            slide_h_pts,
            target_w,
            target_h,
            ppi,
        )

        indices = (list(slide_indices) if slide_indices is not None
                   else list(range(total)))
        selected_total = len(indices)
        for progress_idx, idx in enumerate(indices):
            if cancel_event and cancel_event.is_set():
                logger.info("Export cancelled before slide %d", idx + 1)
                raise InterruptedError("Export cancelled by user.")

            slide_num = idx + 1
            if progress_callback:
                progress_callback(progress_idx, selected_total)

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

            # Step 2+3 — copy all shapes to clipboard, save as PNG
            exported = False
            try:
                shape_range = slide.Shapes.Range()
                shape_range.Copy()
                time.sleep(0.2)
                if _save_clipboard_as_png(out_path, target_w, target_h):
                    logger.info("Exported slide %d → %s (clipboard, %dx%d px)",
                                slide_num, out_path, target_w, target_h)
                    exported = True
                else:
                    logger.debug("Slide %d: no image on clipboard", slide_num)
            except Exception as exc:
                logger.debug("Clipboard approach failed for slide %d: %s",
                             slide_num, exc)

            # Fallback — ShapeRange.Export (may lose transparency)
            if not exported:
                try:
                    shape_range = slide.Shapes.Range()
                    shape_range.Export(out_path, 2, target_w, target_h)
                    logger.info("Exported slide %d → %s (Export fallback, %dx%d px)",
                                slide_num, out_path, target_w, target_h)
                except Exception as exc:
                    logger.warning(
                        "ShapeRange.Export failed for slide %d: %s "
                        "— falling back to slide export",
                        slide_num, exc,
                    )
                    slide.Export(out_path, "PNG", target_w, target_h)
                    logger.info("Exported slide %d → %s (slide fallback)",
                                slide_num, out_path)

            # Step 4 — remove bounding rectangle
            try:
                for i in range(1, slide.Shapes.Count + 1):
                    if slide.Shapes(i).Id == bounding_id:
                        slide.Shapes(i).Delete()
                        break
                logger.debug("Removed bounding rect from slide %d", slide_num)
            except Exception as exc:
                logger.warning(
                    "Could not remove bounding rect from slide %d: %s",
                    slide_num, exc,
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
        pythoncom.CoUninitialize()

    if progress_callback:
        progress_callback(selected_total, selected_total)
