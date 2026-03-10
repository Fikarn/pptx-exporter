"""macOS backend: drive Microsoft PowerPoint via AppleScript / osascript.

Workflow:
1. Open the original .pptx in PowerPoint.
2. For each slide:
   a. Add an invisible full-slide rectangle via AppleScript.
   b. Select all shapes (Esc, Esc, Cmd+A).
   c. Copy the selection (Cmd+C) — this captures all objects including
      placeholders as a transparent PNG on the clipboard.
   d. Save the clipboard PNG to disk via NSPasteboard.
3. Close without saving — the original file is never modified.

The clipboard approach is necessary because AppleScript's ``save as picture``
only works on a single shape, and PowerPoint refuses to group placeholder
shapes. Copying a multi-selection preserves transparency and includes all
objects.
"""

import logging
import shutil
import subprocess
from pathlib import Path
from typing import Callable, Optional

from ..utils import slide_output_name

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# AppleScript templates
# ---------------------------------------------------------------------------

_SCRIPT_OPEN = """\
tell application "Microsoft PowerPoint"
    activate
    open POSIX file "{pptx_path}"
    delay 2
end tell
"""

_SCRIPT_GO_TO_SLIDE = """\
tell application "Microsoft PowerPoint"
    activate
    set theView to view of active window
    set slide of theView to slide {slide_num} of active presentation
    delay 0.5
end tell
"""

# Add an invisible full-slide rectangle to the current slide.
_SCRIPT_ADD_BOUNDING_RECT = """\
tell application "Microsoft PowerPoint"
    set oSlide to slide {slide_num} of active presentation
    set newShape to make new shape at oSlide with properties {{auto shape type:autoshape rectangle, left position:0, top:0, width:{slide_w}, height:{slide_h}}}
    set visible of fill format of newShape to false
    set line weight of line format of newShape to 0
    set transparency of line format of newShape to 1.0
    return "ok"
end tell
"""

# Select all shapes on the current slide and copy to clipboard.
_SCRIPT_SELECT_ALL_AND_COPY = """\
tell application "Microsoft PowerPoint" to activate
delay 0.3
tell application "System Events"
    tell process "Microsoft PowerPoint"
        key code 53
        delay 0.2
        key code 53
        delay 0.2
        keystroke "a" using {command down}
        delay 0.3
        keystroke "c" using {command down}
        delay 0.5
    end tell
end tell
"""

# Save clipboard image data to a high-resolution PNG via NSPasteboard (Cocoa).
# Tries PDF (vector) first for best quality, then TIFF, then PNG.
# Renders into a bitmap at the target pixel dimensions for 300 PPI output.
_SCRIPT_CLIPBOARD_TO_PNG = """\
use framework "AppKit"
use scripting additions

set pb to current application's NSPasteboard's generalPasteboard()
set targetW to {target_w} as integer
set targetH to {target_h} as integer

-- Collect clipboard data: prefer PDF (vector), then TIFF, then PNG
set srcData to missing value
set pdfType to current application's NSPasteboardTypePDF
set srcData to pb's dataForType:pdfType
if srcData is missing value then
    set tiffType to current application's NSPasteboardTypeTIFF
    set srcData to pb's dataForType:tiffType
end if
if srcData is missing value then
    set pngType to current application's NSPasteboardTypePNG
    set srcData to pb's dataForType:pngType
end if
if srcData is missing value then return "no_image"

-- Build NSImage from clipboard data
set img to current application's NSImage's alloc()'s initWithData:srcData
if img is missing value then return "no_image"

-- Create a bitmap at the target resolution
set bitmapRep to (current application's NSBitmapImageRep's alloc()'s initWithBitmapDataPlanes:(missing value) pixelsWide:targetW pixelsHigh:targetH bitsPerSample:8 samplesPerPixel:4 hasAlpha:true isPlanar:false colorSpaceName:(current application's NSCalibratedRGBColorSpace) bytesPerRow:0 bitsPerPixel:0)
bitmapRep's setSize:(current application's NSMakeSize(targetW, targetH))

-- Render the image into the bitmap at target size
current application's NSGraphicsContext's saveGraphicsState()
set ctx to (current application's NSGraphicsContext's graphicsContextWithBitmapImageRep:bitmapRep)
current application's NSGraphicsContext's setCurrentContext:ctx
ctx's setImageInterpolation:(current application's NSImageInterpolationHigh)
img's drawInRect:(current application's NSMakeRect(0, 0, targetW, targetH))
current application's NSGraphicsContext's restoreGraphicsState()

-- Save as PNG
set pngData to (bitmapRep's representationUsingType:(current application's NSBitmapImageFileTypePNG) |properties|:(missing value))
if pngData is missing value then return "convert_failed"

set writeResult to (pngData's writeToFile:"{out_path}" atomically:true) as boolean
if writeResult then return "ok"
return "write_failed"
"""

_SCRIPT_CLOSE_WITHOUT_SAVING = """\
tell application "Microsoft PowerPoint"
    close active presentation saving no
end tell
"""


def _escape_applescript_string(s: str) -> str:
    """Escape a string for safe embedding in AppleScript double-quoted literals."""
    return s.replace("\\", "\\\\").replace('"', '\\"')


def _run_applescript(script: str) -> str:
    """Execute *script* via osascript and return stdout stripped."""
    result = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"AppleScript error (exit {result.returncode}):\n"
            f"STDOUT: {result.stdout.strip()}\n"
            f"STDERR: {result.stderr.strip()}"
        )
    return result.stdout.strip()


def export_slides(
    pptx_path: Path,
    output_dir: Path,
    progress_callback: Optional[Callable[[int, int], None]] = None,
) -> None:
    """Export every slide of *pptx_path* as a transparent PNG into *output_dir*.

    Opens the original file in PowerPoint. For each slide: adds an invisible
    bounding rectangle, selects all, copies to clipboard, and saves the
    clipboard PNG via NSPasteboard. Closes without saving.
    """
    # Read slide metadata via python-pptx (no modification).
    from pptx import Presentation
    prs = Presentation(str(pptx_path))
    total = len(prs.slides)
    slide_w = round(prs.slide_width / 12700, 1)   # EMU → points
    slide_h = round(prs.slide_height / 12700, 1)   # EMU → points
    # Target pixel dimensions for 300 PPI output (points are 1/72 inch).
    target_ppi = 300
    target_w = int(round(slide_w / 72 * target_ppi))
    target_h = int(round(slide_h / 72 * target_ppi))
    logger.info("Slide count: %d, dimensions: %.0f x %.0f pts, target: %d x %d px (%d PPI)",
                total, slide_w, slide_h, target_w, target_h, target_ppi)

    # Open the original file in PowerPoint.
    pptx_posix = str(pptx_path)
    logger.info("Opening presentation in PowerPoint: %s", pptx_posix)
    _run_applescript(_SCRIPT_OPEN.format(
        pptx_path=_escape_applescript_string(pptx_posix)
    ))

    try:
        for idx in range(total):
            slide_num = idx + 1
            if progress_callback:
                progress_callback(idx, total)

            logger.debug("Processing slide %d/%d", slide_num, total)

            out_name = slide_output_name(idx, total)
            out_path = output_dir / out_name

            # Step 1: Navigate to the slide.
            _run_applescript(
                _SCRIPT_GO_TO_SLIDE.format(slide_num=slide_num)
            )

            # Step 2: Add invisible full-slide bounding rectangle.
            _run_applescript(
                _SCRIPT_ADD_BOUNDING_RECT.format(
                    slide_num=slide_num,
                    slide_w=slide_w,
                    slide_h=slide_h,
                )
            )

            # Step 3: Select all shapes and copy to clipboard.
            _run_applescript(_SCRIPT_SELECT_ALL_AND_COPY)

            # Step 4: Save clipboard as high-resolution PNG.
            save_result = _run_applescript(
                _SCRIPT_CLIPBOARD_TO_PNG.format(
                    out_path=_escape_applescript_string(str(out_path)),
                    target_w=target_w,
                    target_h=target_h,
                )
            )

            if save_result == "ok":
                logger.info("Exported slide %d → %s", slide_num, out_path)
            elif save_result == "no_image":
                logger.warning("Slide %d: no image data on clipboard", slide_num)
            else:
                logger.warning("Slide %d: clipboard save returned '%s'", slide_num, save_result)

    finally:
        try:
            _run_applescript(_SCRIPT_CLOSE_WITHOUT_SAVING)
            logger.info("Closed presentation without saving")
        except Exception as exc:
            logger.warning("Could not close presentation: %s", exc)

    if progress_callback:
        progress_callback(total, total)
