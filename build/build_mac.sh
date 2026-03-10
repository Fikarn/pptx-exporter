#!/usr/bin/env bash
# Build a macOS .app bundle with PyInstaller.
# Output: dist/mac/pptx-exporter.app
set -euo pipefail

pip install --upgrade pyinstaller
pip install -e .

pyinstaller \
  --windowed \
  --name "pptx-exporter" \
  --collect-all customtkinter \
  --collect-all darkdetect \
  --collect-data tkinterdnd2 \
  --add-data "vendor:vendor" \
  --hidden-import "pptx_exporter.platforms.macos" \
  --hidden-import "pptx_exporter.platforms.windows" \
  --hidden-import "lxml.etree" \
  --hidden-import "lxml._elementpath" \
  --osx-bundle-identifier "com.fikarn.pptx-exporter" \
  --noconfirm \
  run.py

mkdir -p dist/mac
mv dist/pptx-exporter.app dist/mac/

# Strip quarantine so the app opens without Gatekeeper prompts
xattr -dr com.apple.quarantine dist/mac/pptx-exporter.app 2>/dev/null || true
