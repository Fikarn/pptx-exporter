# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

[Unreleased]: https://github.com/Fikarn/pptx-exporter/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/Fikarn/pptx-exporter/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/Fikarn/pptx-exporter/releases/tag/v0.1.0
