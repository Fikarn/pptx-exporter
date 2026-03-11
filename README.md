# pptx-exporter

[![CI](https://github.com/Fikarn/pptx-exporter/actions/workflows/ci.yml/badge.svg)](https://github.com/Fikarn/pptx-exporter/actions/workflows/ci.yml)
[![Latest Release](https://img.shields.io/github/v/release/Fikarn/pptx-exporter)](https://github.com/Fikarn/pptx-exporter/releases/latest)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

A Python desktop application that exports the objects from each slide of a PowerPoint file as **transparent PNG images** — exactly as if you had manually selected all objects on each slide and used "Copy" → "Paste as Picture" in PowerPoint.

---

## Why pptx-exporter?

Exporting a PowerPoint slide the normal way (File → Export) always produces an image with the slide background. pptx-exporter instead selects all objects on each slide and exports only them, resulting in PNGs with **fully transparent backgrounds** — ideal for compositing, web use, or design workflows.

---

## How it works

For every slide the app:

1. Opens the original `.pptx` file in Microsoft PowerPoint.
2. Navigates to the slide and adds an invisible, borderless rectangle exactly the size of the slide — a bounding anchor so the exported PNG always has the full slide dimensions.
3. Selects all objects (Cmd+A / Ctrl+A) and copies to the clipboard.
4. Reads the clipboard image — preferring PDF vector data for maximum fidelity, falling back to TIFF or PNG.
5. Renders the clipboard data at the chosen resolution (default 300 dpi, configurable 36–2400) and saves it as a transparent PNG (`slide_01.png`, `slide_02.png`, …).
6. Closes the presentation without saving — the original file is never modified.

---

## Requirements

Microsoft PowerPoint must be installed. The app detects PowerPoint at startup and disables the export button if it is not found.

| OS | Backend | Pre-built binary |
|---|---|---|
| macOS | AppleScript automation | ✅ Available |
| Windows | COM automation (pywin32) | 🚧 Coming soon |

---

## Download (macOS)

Download the latest `pptx-exporter-macos-vX.Y.Z.dmg` from the [Releases page](https://github.com/Fikarn/pptx-exporter/releases), open the disk image, and drag `pptx-exporter.app` to your Applications folder.

### Bypassing Gatekeeper

Because the app is unsigned, macOS may silently block it. The most reliable fix is to remove the quarantine flag in Terminal:

```bash
xattr -dr com.apple.quarantine /Applications/pptx-exporter.app
```

Then double-click to open normally. Alternatively, right-click → **Open** → **Open**.

---

## Usage

1. Launch the app.
2. **Select file(s)** — click **Browse…** to pick one or more `.pptx` files. Selecting multiple files enables batch export, where each file is exported into its own subfolder.
3. **Choose an output folder** — a default is suggested automatically; click **Browse…** to change it.
4. **Pick a resolution** — choose 72, 150, or 300 dpi from the presets, or select **Custom** to enter any value between 36 and 2400 dpi.
5. **Select slides** *(optional)* — when a single file with multiple slides is selected, uncheck "All slides" and enter a range (e.g. `1-5, 8, 10-12`) to export only specific slides.
6. Click **Export PNGs**.
7. A progress bar tracks each slide. Click **Cancel** at any time to stop cleanly after the current slide finishes.
8. When done, click **Open Folder ↗** to reveal the exported PNGs.

The selected resolution and last-used output folder are remembered across sessions.

---

## Building from source

### Prerequisites

- Python 3.10 or later
- Microsoft PowerPoint installed on the same machine

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

### Run tests

```bash
pytest
```

### Lint

```bash
flake8 src/ tests/
```

### Build a standalone binary

**macOS** — produces `dist/mac/pptx-exporter.app` and `dist/mac/pptx-exporter.dmg`:

```bash
bash build/build_mac.sh
```

**Windows** — produces `dist\windows\pptx-exporter.exe`:

```bat
build\build_windows.bat
```

### Configure log level

```bash
LOG_LEVEL=DEBUG python -m pptx_exporter.main
```

Valid values: `DEBUG`, `INFO`, `WARNING`, `ERROR`, `CRITICAL` (default: `INFO`).

---

## CI/CD

- **CI** runs on every push and pull request to `main`: installs dependencies, runs `pytest`, and runs `flake8`.
- **Release** is triggered by pushing a version tag (e.g. `v1.0.3`): runs tests, builds the macOS `.dmg` and Windows `.exe`, and creates a GitHub Release with both attached. The Windows build runs in parallel but does not block the release if it fails.

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

Follow PEP 8 style and use type hints on all function signatures.

---

## License

[MIT](LICENSE) © Fikarn
