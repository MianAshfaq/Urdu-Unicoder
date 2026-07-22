@echo off
title Urdu Unicoder
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo The software is not installed yet.
    echo Running setup first...
    call setup_windows.bat
)

".venv\Scripts\python.exe" app\main.py
if errorlevel 1 (
    echo.
    echo The application closed with an error.
    pause
)
