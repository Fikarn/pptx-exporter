@echo off
REM Build a single-file Windows .exe with PyInstaller.
REM Output: dist\windows\pptx-exporter.exe

pip install --upgrade pyinstaller
pip install -e ".[windows]"

pyinstaller ^
  --onefile ^
  --windowed ^
  --name pptx-exporter ^
  --collect-all customtkinter ^
  --collect-all darkdetect ^
  --hidden-import "pptx_exporter.platforms.macos" ^
  --hidden-import "pptx_exporter.platforms.windows" ^
  --hidden-import "lxml.etree" ^
  --hidden-import "lxml._elementpath" ^
  --noconfirm ^
  run.py

if not exist dist\windows mkdir dist\windows
move dist\pptx-exporter.exe dist\windows\
