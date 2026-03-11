# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project does

pptx-exporter is a Python desktop app that exports each PowerPoint slide's objects as **transparent PNG images** — removing the slide background. It drives Microsoft PowerPoint itself via platform-native automation (AppleScript on macOS, COM/pywin32 on Windows) to select all objects, copy them, and save the clipboard as a high-resolution PNG.

## Commands

```bash
# Run the app
python -m pptx_exporter.main

# Run all tests
pytest

# Run a single test file or test
pytest tests/test_utils.py
pytest tests/test_exporter.py::test_full_export_flow

# Lint
flake8 src/ tests/

# Build macOS .app
bash build/build_mac.sh

# Build Windows .exe
build\build_windows.bat

# Debug logging
LOG_LEVEL=DEBUG python -m pptx_exporter.main
```

Setup: `pip install -r requirements-dev.txt && pip install -e .` (use `pip install -e ".[windows]"` on Windows for pywin32).

## Architecture

The app has three layers: **GUI** → **Exporter** → **Platform backend**.

- `gui.py` — CustomTkinter UI. Runs export in a background `threading.Thread` with cancel support via `threading.Event`. Progress updates flow back to the main thread via `self.after()`.
- `exporter.py` — `Exporter` class. Detects the backend at init (`detect_backend()`), validates inputs, dispatches to the correct platform module. This is the only public API for exports.
- `platforms/macos.py` — AppleScript backend. Opens the file in PowerPoint, per slide: adds an invisible bounding rectangle, Esc+Esc+Cmd+A+Cmd+C via System Events, reads clipboard via NSPasteboard (prefers PDF vector data → renders into NSBitmapImageRep at target PPI), saves as PNG. Closes without saving. Checks Accessibility permissions before starting. Uses polling (not fixed delays) to wait for PowerPoint readiness, and retries clipboard reads up to 3 times.
- `platforms/windows.py` — COM backend via win32com. Per slide: adds bounding rectangle, copies all shapes to clipboard via `ShapeRange.Copy()`, reads clipboard (prefers PNG format, falls back to CF_DIB with alpha), resizes to target PPI via Pillow, saves as PNG. Falls back to `ShapeRange.Export` if clipboard fails. Removes bounding rectangle after each slide. Uses `pythoncom.CoInitialize()` for thread safety.
- `utils.py` — Shared helpers: path validation, OS/backend detection, logging config, slide filename generation.

### Key design decisions

- **PowerPoint is required** — there is no pure-Python fallback. The GUI disables the export button if PowerPoint is not detected.
- **Bounding rectangle trick** — an invisible full-slide rectangle is added before selecting all shapes. This ensures the exported PNG has exact slide dimensions even if objects don't span the full slide.
- **Both platforms use clipboard** — because AppleScript's `save as picture` only works on single shapes, and PowerPoint refuses to group placeholder shapes. The clipboard approach preserves transparency and allows rendering at any target PPI. Windows falls back to `ShapeRange.Export` if clipboard fails.
- **Coordinate math**: EMU → points via `/ 12700`, points → pixels via `/ 72 * PPI`.

## Code style

- PEP 8 with max line length 99 (configured in `.flake8` and `pyproject.toml`).
- Type hints on all function signatures.
- `src/` layout — package is at `src/pptx_exporter/`, entry point for PyInstaller is `run.py` at repo root.

## Release process

Push a version tag: `git tag v1.2.3 && git push origin v1.2.3`. This triggers the release workflow which runs tests, builds macOS .app and Windows .exe, and creates a GitHub Release. The Windows build is best-effort (`continue-on-error: true`).
