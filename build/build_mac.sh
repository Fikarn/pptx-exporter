#!/usr/bin/env bash
# Build a macOS .app bundle with PyInstaller and package it as a .dmg.
# Output: dist/mac/pptx-exporter.app
#         dist/mac/pptx-exporter.dmg
set -euo pipefail

pip install --upgrade pyinstaller
pip install -e .

pyinstaller \
  --windowed \
  --name "pptx-exporter" \
  --collect-all customtkinter \
  --collect-all darkdetect \
  --hidden-import "pptx_exporter.gui.app" \
  --hidden-import "pptx_exporter.platforms.macos" \
  --hidden-import "pptx_exporter.platforms.windows" \
  --hidden-import "lxml.etree" \
  --hidden-import "lxml._elementpath" \
  --add-data "src/pptx_exporter/tkdnd/macos-arm64:pptx_exporter/tkdnd/macos-arm64" \
  --add-data "src/pptx_exporter/tkdnd/macos-x86_64:pptx_exporter/tkdnd/macos-x86_64" \
  --osx-bundle-identifier "com.fikarn.pptx-exporter" \
  --noconfirm \
  run.py

mkdir -p dist/mac
mv dist/pptx-exporter.app dist/mac/

# Strip quarantine so the app opens without Gatekeeper prompts
xattr -dr com.apple.quarantine dist/mac/pptx-exporter.app 2>/dev/null || true

# ── Create .dmg disk image ────────────────────────────────────────────────
# Stage the .app and an Applications symlink in a temp folder, then use
# hdiutil to create a compressed read-only DMG (no external tools needed).
DMG_STAGING="$(mktemp -d)"
cp -R dist/mac/pptx-exporter.app "$DMG_STAGING/"
ln -s /Applications "$DMG_STAGING/Applications"

hdiutil create \
  -volname "pptx-exporter" \
  -srcfolder "$DMG_STAGING" \
  -ov \
  -format UDZO \
  dist/mac/pptx-exporter.dmg

rm -rf "$DMG_STAGING"
echo "DMG created: dist/mac/pptx-exporter.dmg"
