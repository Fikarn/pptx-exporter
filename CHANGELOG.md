# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

[Unreleased]: https://github.com/Fikarn/pptx-exporter/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/Fikarn/pptx-exporter/releases/tag/v0.1.0
