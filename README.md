# pptx-exporter

[![CI](https://github.com/Fikarn/pptx-exporter/actions/workflows/ci.yml/badge.svg)](https://github.com/Fikarn/pptx-exporter/actions/workflows/ci.yml)
[![Latest Release](https://img.shields.io/github/v/release/Fikarn/pptx-exporter)](https://github.com/Fikarn/pptx-exporter/releases/latest)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

A Python desktop application that automates exporting the objects from each slide of a PowerPoint file as **transparent PNG images** — exactly as if you had manually selected all objects on each slide and used "Copy" → "Paste as Picture" in PowerPoint.

---

## Why pptx-exporter?

Exporting a PowerPoint slide as a whole (File → Export) always produces an image with the slide background. pptx-exporter instead selects all objects on each slide and exports only them, resulting in PNGs with **fully transparent backgrounds** — ideal for compositing, web use, or design workflows.

---

## How it works

For every slide, the app:

1. Opens the original `.pptx` file in Microsoft PowerPoint.
2. Navigates to the slide and adds an invisible, transparent, borderless rectangle exactly the size of the slide. This acts as a bounding anchor so the exported PNG always has the full slide dimensions, regardless of where individual objects sit.
3. Selects all objects (Cmd+A / Ctrl+A) and copies to the clipboard (Cmd+C / Ctrl+C).
4. Reads the clipboard image — preferring PDF vector data for maximum fidelity, falling back to TIFF or PNG.
5. Renders the clipboard data into a **300 PPI** bitmap (e.g. 4000×2250 px for a standard 16:9 slide) and saves it as a transparent PNG (`slide_01.png`, `slide_02.png`, …).
6. Closes the presentation without saving — the original file is never modified.

---

## Requirements

Microsoft PowerPoint must be installed. The app detects PowerPoint at startup and disables the export button if it is not found.

| OS | Backend |
|---|---|
| macOS | AppleScript (osascript) — 300 PPI, full transparency support |
| Windows | COM automation (pywin32) — full transparency support |

The active backend is displayed in the app's status banner when it launches.

---

## Releases — download the app

Pre-built binaries are available on the [GitHub Releases page](https://github.com/Fikarn/pptx-exporter/releases):

- **macOS:** Download `pptx-exporter-macos-vX.Y.Z.zip`, unzip, and open `pptx-exporter.app`.
- **Windows:** Download `pptx-exporter-windows-vX.Y.Z.zip`, unzip, and run `pptx-exporter.exe`.

No Python installation required.

### macOS — bypassing Gatekeeper

Because the app is not signed with an Apple Developer certificate, macOS Gatekeeper will block it on first launch. To open it:

1. Right-click (or Control-click) `pptx-exporter.app` in Finder.
2. Select **Open**.
3. Click **Open** in the dialog that appears.

You only need to do this once. After that, you can double-click to open normally.

### Windows — bypassing SmartScreen

Because the `.exe` is not signed with a code-signing certificate, Windows SmartScreen may warn you. To run it:

1. Click **More info** in the SmartScreen dialog.
2. Click **Run anyway**.

For production distribution, consider purchasing a code-signing certificate.

---

## Usage

1. Launch the app (see Releases above, or run from source — see below).
2. Click **Browse…** next to "Input .pptx" and select your PowerPoint file.
3. Click **Browse…** next to "Output folder" and select or create a destination folder.
4. Click **Run Export**.
5. The progress bar shows which slide is being processed. When complete, a dialog confirms the output location.

---

## Building from source

### Prerequisites

- Python 3.10 or later
- `pip`

### Install

```bash
git clone https://github.com/Fikarn/pptx-exporter.git
cd pptx-exporter
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements-dev.txt
pip install -e .
```

### Run

```bash
python -m pptx_exporter.main
```

Or, after installing with `pip install -e .`:

```bash
pptx-exporter
```

### Run tests

```bash
pytest
```

### Lint

```bash
flake8 src/ tests/
```

### Build a standalone binary

**macOS:**

```bash
bash build/build_mac.sh
# Output: dist/mac/pptx-exporter.app
```

**Windows:**

```bat
build\build_windows.bat
REM Output: dist\windows\pptx-exporter.exe
```

### Configure log level

Set the `LOG_LEVEL` environment variable before running:

```bash
LOG_LEVEL=DEBUG pptx-exporter
```

Valid values: `DEBUG`, `INFO`, `WARNING`, `ERROR`, `CRITICAL` (default: `INFO`).

---

## CI/CD

- **CI** runs on every push and pull request to `main`: installs dependencies, runs `pytest`, and runs `flake8`.
- **Release** is triggered by pushing a version tag (e.g. `v1.2.3`): runs tests, builds macOS `.app` and Windows `.exe` in parallel, then creates a GitHub Release with both archives attached.

To publish a new release:

```bash
git tag v1.2.3
git push origin v1.2.3
```

---

## Contributing

Contributions are welcome. Please:

1. Fork the repository and create a feature branch.
2. Write tests for new functionality.
3. Ensure `pytest` and `flake8` pass.
4. Open a pull request against `main`.

Please follow PEP 8 style and use type hints on all function signatures.

---

## License

[MIT](LICENSE) © Fikarn
