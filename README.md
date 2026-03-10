# pptx-exporter

[![CI](https://github.com/Fikarn/pptx-exporter/actions/workflows/ci.yml/badge.svg)](https://github.com/Fikarn/pptx-exporter/actions/workflows/ci.yml)
[![Latest Release](https://img.shields.io/github/v/release/Fikarn/pptx-exporter)](https://github.com/Fikarn/pptx-exporter/releases/latest)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

A Python desktop application that automates exporting the objects from each slide of a PowerPoint file as **transparent PNG images** — exactly as if you had manually selected all objects on each slide and used "Save as Picture" in PowerPoint.

---

## Why pptx-exporter?

Exporting a PowerPoint slide as a whole (File → Export) always produces an image with the slide background. pptx-exporter instead selects all objects on each slide and exports only them, resulting in PNGs with **fully transparent backgrounds** — ideal for compositing, web use, or design workflows.

---

## How it works

For every slide, the app:

1. Adds an invisible, transparent, borderless rectangle exactly the size of the slide. This acts as a bounding anchor so the exported PNG is always the full slide dimensions, regardless of where individual objects sit.
2. Selects all objects on the slide.
3. Exports the selection as a transparent PNG (`slide_01.png`, `slide_02.png`, …).
4. Removes the bounding rectangle, leaving the original file unmodified.

---

## Backends

pptx-exporter detects your OS and whether Microsoft PowerPoint is installed at startup, and picks the best available backend:

| OS | PowerPoint installed? | Backend used |
|---|---|---|
| macOS | Yes | AppleScript (osascript) — full transparency support |
| macOS | No | python-pptx + Pillow (picture shapes only) |
| Windows | Yes | COM automation (pywin32) — full transparency support |
| Windows | No | python-pptx + Pillow (picture shapes only) |

The active backend is displayed in the app's status banner when it launches.

> **Note:** The fallback backend (python-pptx + Pillow) can only render embedded picture/image shapes. Text boxes, auto-shapes, and other vector elements are skipped with a warning. For full fidelity, use the app with Microsoft PowerPoint installed.

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
