@echo off
title Build Urdu Unicoder EXE
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    call setup_windows.bat
)

echo Installing PyInstaller...
".venv\Scripts\python.exe" -m pip install pyinstaller

echo Building Windows executable...
".venv\Scripts\python.exe" -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --name "Urdu Unicoder" ^
  --icon "assets\urdu-unicoder.ico" ^
  --add-data "assets;assets" ^
  --collect-all PySide6 ^
  --collect-all fitz ^
  --collect-all docx ^
  app\main.py

echo.
echo Build complete.
echo Check the dist folder.
pause
