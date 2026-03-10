"""macOS backend: drive Microsoft PowerPoint via AppleScript / osascript.

Per-slide workflow (mirrors manual user actions):
1. Add a transparent, borderless full-slide rectangle via AppleScript.
2. Select all objects on the slide (cmd+a equivalent).
3. Export the selection as a transparent PNG via "Save as Picture".
4. Remove the bounding rectangle.
"""

import logging
import subprocess
import tempfile
from pathlib import Path
from typing import Callable, Optional

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# AppleScript templates
# ---------------------------------------------------------------------------

_SCRIPT_OPEN_PRESENTATION = """\
tell application "Microsoft PowerPoint"
    set theDoc to open POSIX file "{pptx_path}"
    set thePath to POSIX path of (path of theDoc)
    return thePath
end tell
"""

_SCRIPT_SLIDE_COUNT = """\
tell application "Microsoft PowerPoint"
    tell document "{doc_name}"
        return count of slides
    end tell
end tell
"""

_SCRIPT_SLIDE_DIMENSIONS = """\
tell application "Microsoft PowerPoint"
    tell document "{doc_name}"
        set w to slide width
        set h to slide height
        return (w as string) & "," & (h as string)
    end tell
end tell
"""

_SCRIPT_ADD_BOUNDING_RECT = """\
tell application "Microsoft PowerPoint"
    tell document "{doc_name}"
        tell slide {slide_num}
            set theShape to make new shape at end of shapes with properties {{\\
                shape type:auto shape, \\
                auto shape type:rounded rectangle, \\
                left position:0, \\
                top:0, \\
                width:{width}, \\
                height:{height}}}
            tell theShape
                tell fill
                    set visible to false
                end tell
                tell line format
                    set line style to no line
                end tell
            end tell
            return (id of theShape) as string
        end tell
    end tell
end tell
"""

_SCRIPT_SELECT_ALL_AND_EXPORT = """\
tell application "Microsoft PowerPoint"
    tell document "{doc_name}"
        tell slide {slide_num}
            -- select all shapes
            set sel to {}
            repeat with s in shapes
                set end of sel to s
            end repeat
            if (count of sel) = 0 then return "no_shapes"
            select (item 1 of sel)
            repeat with i from 2 to count of sel
                tell application "System Events"
                    key code 0 using command down
                end tell
            end repeat
            -- export as picture
            save as picture (item 1 of sel) in "{out_path}" \\
                as save as PNG
        end tell
    end tell
end tell
"""

_SCRIPT_EXPORT_SLIDE_SHAPES = """\
tell application "Microsoft PowerPoint"
    tell document "{doc_name}"
        tell slide {slide_num}
            set shapeList to every shape
            if (count of shapeList) is 0 then return "no_shapes"
            tell shapeList to select
            save selection of document "{doc_name}" in "{out_path}" as save as PNG
        end tell
    end tell
end tell
"""

_SCRIPT_REMOVE_SHAPE_BY_ID = """\
tell application "Microsoft PowerPoint"
    tell document "{doc_name}"
        tell slide {slide_num}
            delete (first shape whose id is {shape_id})
        end tell
    end tell
end tell
"""

_SCRIPT_CLOSE_WITHOUT_SAVING = """\
tell application "Microsoft PowerPoint"
    close document "{doc_name}" saving no
end tell
"""


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

    Uses AppleScript to drive Microsoft PowerPoint for Mac.

    Args:
        pptx_path: Resolved path to the .pptx file.
        output_dir: Resolved path to the output directory.
        progress_callback: Optional callable(slide_index, total_slides).
    """
    pptx_posix = str(pptx_path)
    doc_name = pptx_path.name

    logger.info("Opening presentation in PowerPoint: %s", pptx_posix)

    # Open the presentation
    open_script = _SCRIPT_OPEN_PRESENTATION.format(pptx_path=pptx_posix)
    _run_applescript(open_script)

    try:
        # Get slide count
        count_script = _SCRIPT_SLIDE_COUNT.format(doc_name=doc_name)
        total = int(_run_applescript(count_script))
        logger.info("Slide count: %d", total)

        # Get slide dimensions (points)
        dim_script = _SCRIPT_SLIDE_DIMENSIONS.format(doc_name=doc_name)
        dims = _run_applescript(dim_script).split(",")
        slide_w_pts = float(dims[0].strip())
        slide_h_pts = float(dims[1].strip())
        logger.debug("Slide dimensions: %.1f x %.1f pts", slide_w_pts, slide_h_pts)

        width = len(str(total))

        for idx in range(total):
            slide_num = idx + 1
            if progress_callback:
                progress_callback(idx, total)

            logger.debug("Processing slide %d/%d", slide_num, total)

            out_name = f"slide_{slide_num:0{width}d}.png"
            out_path = str(output_dir / out_name)

            # Step 1 — add invisible bounding rectangle
            add_script = _SCRIPT_ADD_BOUNDING_RECT.format(
                doc_name=doc_name,
                slide_num=slide_num,
                width=slide_w_pts,
                height=slide_h_pts,
            )
            shape_id_str = _run_applescript(add_script).strip()
            logger.debug("Added bounding rect, shape id: %s", shape_id_str)

            # Steps 2+3 — select all shapes and export as PNG
            export_script = _SCRIPT_EXPORT_SLIDE_SHAPES.format(
                doc_name=doc_name,
                slide_num=slide_num,
                out_path=out_path,
            )
            result = _run_applescript(export_script)
            if result == "no_shapes":
                logger.warning("Slide %d has no shapes, skipping", slide_num)
            else:
                logger.info("Exported slide %d → %s", slide_num, out_path)

            # Step 4 — remove bounding rectangle
            try:
                shape_id = int(shape_id_str)
                remove_script = _SCRIPT_REMOVE_SHAPE_BY_ID.format(
                    doc_name=doc_name,
                    slide_num=slide_num,
                    shape_id=shape_id,
                )
                _run_applescript(remove_script)
                logger.debug("Removed bounding rect from slide %d", slide_num)
            except Exception as exc:
                logger.warning(
                    "Could not remove bounding rect from slide %d: %s", slide_num, exc
                )

    finally:
        # Always close the document without saving
        try:
            close_script = _SCRIPT_CLOSE_WITHOUT_SAVING.format(doc_name=doc_name)
            _run_applescript(close_script)
            logger.info("Closed presentation without saving")
        except Exception as exc:
            logger.warning("Could not close presentation: %s", exc)

    if progress_callback:
        progress_callback(total, total)
