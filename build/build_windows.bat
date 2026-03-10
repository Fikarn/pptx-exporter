@echo off
REM Build a Windows .exe with PyInstaller.
REM Output: dist\windows\pptx-exporter.exe

pip install --upgrade pyinstaller
pip install -e .

pyinstaller ^
  --windowed ^
  --name pptx-exporter ^
  --collect-data customtkinter ^
  --collect-data tkinterdnd2 ^
  --add-data "vendor;vendor" ^
  --noconfirm ^
  src\pptx_exporter\main.py

if not exist dist\windows mkdir dist\windows
move dist\pptx-exporter.exe dist\windows\
