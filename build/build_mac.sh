#!/usr/bin/env bash
# Build a macOS .app bundle with PyInstaller.
# Output: dist/mac/pptx-exporter.app
set -euo pipefail

pip install --upgrade pyinstaller
pip install -e .

pyinstaller \
  --windowed \
  --name "pptx-exporter" \
  --collect-data customtkinter \
  --collect-data tkinterdnd2 \
  --add-data "vendor:vendor" \
  --noconfirm \
  src/pptx_exporter/main.py

mkdir -p dist/mac
mv dist/pptx-exporter.app dist/mac/
