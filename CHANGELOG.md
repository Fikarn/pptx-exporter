# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.3.0] - 2026-03-11

### Changed
- **Complete GUI redesign** ("Quiet Utility"): new layout splits the interface into a File Card, Settings Card, and Action Area, replacing the previous flat vertical form.
- **File Card**: large drop zone is the primary empty state; switches to a compact file list with icon, slide count, and remove button per file. Scrollable via `CTkScrollableFrame` when more than 4 files are loaded.
- **Settings Card**: Resolution, Slides, and Output rows consolidated into a single rounded card with internal dividers. Resolution presets now show as `72 / 150 / 300 / Custom` labels (previously `72 dpi` etc.).
- **Action Area**: Export button is taller (44px) with a disabled-state hint line ("Select a file to begin" / "PowerPoint not found"). Progress bar is slimmer (4px). Open Folder button appears inline below the progress bar after a successful export.
- **Error banner**: redesigned with a red left-border accent instead of a full red background; dismiss button is an ✕ icon.
- **Inline validation**: slide range errors and duplicate-file warnings now appear inline instead of `messagebox.showwarning()` popups.
- **Keyboard shortcuts**: Cmd+O / Ctrl+O opens the file browser; Escape cancels an in-progress export.
- **Code structure**: `gui.py` (1028 lines) replaced by a `gui/` package — `app.py`, `tokens.py`, `settings.py`, and `widgets/` sub-package with one file per widget class.
- Build scripts updated with `--hidden-import pptx_exporter.gui.app`.

## [1.2.0] - 2026-03-11

### Added
- **Drag-and-drop**: drop `.pptx` files onto the file panel to add them. Uses vendored tkdnd 2.9.5 binaries with Tcl 9 support (macOS) and Tcl 8.6 (Windows). Falls back gracefully to browse-only if tkdnd cannot load.
- **Improved batch UX**: file panel now shows a scrollable list of loaded files with per-file remove buttons, an "Add files" button to append more files, and "Clear all" to reset. Browsing and dropping always appends to the existing list instead of replacing it.
- Duplicate file detection: warns when the same file is added twice.
- Non-`.pptx` file rejection: warns and skips unsupported files when dropped.
- Visual drop feedback: file panel border highlights on drag-over.

### Changed
- Section label renamed from "INPUT FILE" to "INPUT FILES".
- `_FileCard` widget replaced by `_FilePanel` with richer file list display.

## [1.1.0] - 2026-03-11

### Added
- **Batch export**: select multiple `.pptx` files at once; each is exported into its own subfolder with aggregate progress tracking.
- **Per-slide selection**: uncheck "All slides" and enter a range (e.g. `1-5, 8, 10-12`) to export only specific slides.
- **Custom PPI**: enter any resolution between 36 and 2400 dpi in addition to the 72/150/300 presets.
- **macOS .dmg packaging**: the macOS build now produces a DMG disk image with an Applications symlink for drag-to-install.
- **Accessibility pre-flight check** on macOS: warns immediately if System Events access is not granted, instead of silently producing empty exports.
- **Overwrite warning**: prompts before exporting into a folder that already contains slide PNGs.
- Partial export indication: error messages now report how many slides were exported before a failure.
- CLAUDE.md for repository context.

### Changed
- **Windows backend rewritten** to use clipboard as the primary export method (PNG format preferred, CF_DIB 32bpp BGRA fallback), with a fallback chain to `ShapeRange.Export` and `slide.Export`.
- Windows backend now computes target pixel dimensions from PPI (was previously ignored).
- Windows backend adds `pythoncom.CoInitialize()`/`CoUninitialize()` for COM thread safety.
- macOS backend replaces hardcoded AppleScript delays with polling loops for PowerPoint readiness.
- macOS backend retries clipboard read up to 3 times per slide to handle timing races.
- CI now runs tests on macOS and Windows runners (previously Ubuntu only); lint separated into its own job.
- Release workflow publishes `.dmg` for macOS instead of `.zip`.
- Dynamic version via `importlib.metadata` instead of hardcoded `__init__.__version__`.

### Removed
- `platforms/fallback.py` (dead code since v0.2.0).
- tkinterdnd2 dependency and vendored tkdnd binary (drag-and-drop broken on macOS/Tcl 9).

## [1.0.3] - 2026-03-10

### Fixed
- PyInstaller build: use a top-level `run.py` launcher instead of
  `src/pptx_exporter/main.py` as the entry point, fixing the
  "attempted relative import with no known parent package" crash on launch.

## [1.0.2] - 2026-03-10

### Fixed
- macOS build: use `--collect-all` for customtkinter and darkdetect to ensure
  theme assets are fully included; add hidden imports for lxml and platform
  backends so PyInstaller does not miss dynamically imported modules.
- macOS build: strip Gatekeeper quarantine flag from the .app bundle during
  build so downloaded apps open without being silently blocked.
- README: document the `xattr` quarantine workaround for macOS users.

## [1.0.1] - 2026-03-10

### Fixed
- Release workflow: Windows build failure no longer blocks the macOS release from publishing.
- README: corrected usage instructions, resolution options, and binary availability.

## [1.0.0] - 2026-03-10

### Added
- Modern GUI redesign built on CustomTkinter with system light/dark mode support.
- Drag-and-drop infrastructure (tkdnd 2.9.5 vendored binary built for Tcl/Tk 9.0; later removed due to Tcl 9 incompatibility).
- Cancel button: cleanly aborts an in-progress export between slides.
- Open Folder button: appears after a successful export to reveal the output directory.
- File metadata card: shows slide count, file size, and output pixel dimensions after selecting a file.
- Configurable export resolution: 72 / 150 / 300 dpi segmented control (default 300).
- Settings persistence: selected resolution and last-used output folder are saved across sessions (`~/.pptx-exporter-settings.json`).
- Smart default output path: automatically suggests `{filename}_pngs/` next to the source file.
- Inline error banner replacing disruptive modal dialogs.
- Backend status pill in the header (green = PowerPoint ready, red = not found).

## [0.2.0] - 2026-03-10

### Changed
- macOS backend rewritten to use the clipboard approach instead of AppleScript's `save as picture`. PowerPoint now copies the selection (Cmd+C) and the image is read from NSPasteboard, preferring PDF vector data for maximum quality.
- macOS backend now exports at **300 PPI** (e.g. 4000×2250 px for a standard 16:9 slide), rendered via NSBitmapImageRep.
- Microsoft PowerPoint is now **required** on both macOS and Windows. The app disables the export button and shows an error message if PowerPoint is not detected.

### Removed
- Fallback backend (python-pptx + Pillow): removed. Full PowerPoint automation is now mandatory for correct transparency and fidelity.

## [0.1.0] - 2026-03-10

### Added
- Initial release of pptx-exporter.
- Tkinter GUI with input file and output folder selectors, progress bar, and status label.
- macOS backend: drives Microsoft PowerPoint via AppleScript (osascript) to export per-slide shape selections as transparent PNG images.
- Windows backend: drives Microsoft PowerPoint via COM automation (pywin32) to export per-slide shape selections as transparent PNG images.
- Fallback backend: uses python-pptx + Pillow to composite picture shapes onto a transparent RGBA canvas when PowerPoint is not installed.
- Per-slide bounding rectangle workflow: adds an invisible full-slide rectangle before export and removes it afterwards, ensuring consistent canvas dimensions.
- Automatic OS and backend detection at startup with a descriptive status banner in the UI.
- `src/` layout with clean module separation: `gui`, `exporter`, `platforms/macos`, `platforms/windows`, `platforms/fallback`, `utils`.
- pytest test suite for `exporter` and `utils` with mocked OS/PowerPoint calls.
- CI workflow (GitHub Actions): runs tests and flake8 on every push and PR to main.
- Release workflow (GitHub Actions): builds macOS `.app` and Windows `.exe` in parallel on version tags, creates a GitHub Release with zip attachments and changelog section as body.
- PyInstaller build scripts for macOS (`build/build_mac.sh`) and Windows (`build/build_windows.bat`).
- `pyproject.toml` following PEP 517/518; version is the single source of truth.
- MIT License.

[Unreleased]: https://github.com/Fikarn/pptx-exporter/compare/v1.3.0...HEAD
[1.3.0]: https://github.com/Fikarn/pptx-exporter/compare/v1.2.0...v1.3.0
[1.2.0]: https://github.com/Fikarn/pptx-exporter/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/Fikarn/pptx-exporter/compare/v1.0.3...v1.1.0
[1.0.3]: https://github.com/Fikarn/pptx-exporter/compare/v1.0.2...v1.0.3
[1.0.2]: https://github.com/Fikarn/pptx-exporter/compare/v1.0.1...v1.0.2
[1.0.1]: https://github.com/Fikarn/pptx-exporter/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/Fikarn/pptx-exporter/compare/v0.2.0...v1.0.0
[0.2.0]: https://github.com/Fikarn/pptx-exporter/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/Fikarn/pptx-exporter/releases/tag/v0.1.0
