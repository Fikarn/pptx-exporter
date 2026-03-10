#!/usr/bin/env bash
# build/build_mac.sh — Build pptx-exporter.app for macOS using PyInstaller
# Usage: bash build/build_mac.sh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="$(dirname "$SCRIPT_DIR")"
cd "$ROOT_DIR"

echo "==> Installing dependencies…"
pip install --upgrade pip
pip install -r requirements-dev.txt

echo "==> Running tests…"
pytest tests/ -x -q

echo "==> Running flake8…"
flake8 src/ tests/

echo "==> Building macOS .app with PyInstaller…"
pyinstaller \
    --name "pptx-exporter" \
    --windowed \
    --onedir \
    --clean \
    --noconfirm \
    --distpath dist/mac \
    --add-data "src/pptx_exporter:pptx_exporter" \
    src/pptx_exporter/main.py

echo ""
echo "==> Build complete: dist/mac/pptx-exporter.app"
echo ""
echo "NOTE: To run the app on macOS without Gatekeeper blocking it,"
echo "right-click the .app in Finder and select 'Open', then confirm."
echo "(This is only needed the first time for unsigned builds.)"
