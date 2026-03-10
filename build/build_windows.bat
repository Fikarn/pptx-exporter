@echo off
REM build\build_windows.bat — Build pptx-exporter.exe for Windows using PyInstaller
REM Usage: build\build_windows.bat

setlocal enabledelayedexpansion

echo =^> Installing dependencies...
pip install --upgrade pip
pip install -r requirements-dev.txt
pip install pywin32

echo =^> Running tests...
pytest tests\ -x -q
if errorlevel 1 (
    echo Tests failed — aborting build.
    exit /b 1
)

echo =^> Running flake8...
flake8 src\ tests\
if errorlevel 1 (
    echo Flake8 violations found — aborting build.
    exit /b 1
)

echo =^> Building Windows .exe with PyInstaller...
pyinstaller ^
    --name "pptx-exporter" ^
    --windowed ^
    --onefile ^
    --clean ^
    --noconfirm ^
    --distpath dist\windows ^
    --add-data "src\pptx_exporter;pptx_exporter" ^
    src\pptx_exporter\main.py

echo.
echo =^> Build complete: dist\windows\pptx-exporter.exe
echo.
echo NOTE: Windows SmartScreen may warn when running an unsigned .exe.
echo Click "More info" then "Run anyway" to proceed.
echo (Consider code-signing with a certificate for production distribution.)
