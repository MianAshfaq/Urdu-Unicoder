@echo off
title Urdu Unicoder Setup
cd /d "%~dp0"

echo ==========================================
echo Urdu Unicoder - Windows Setup
echo ==========================================
echo.

where py >nul 2>nul
if errorlevel 1 (
    echo Python Launcher was not found.
    echo Install Python 3.11, 3.12, or 3.13 from python.org.
    echo During installation, enable "Add Python to PATH".
    pause
    exit /b 1
)

echo Creating virtual environment...
py -3 -m venv .venv
if errorlevel 1 (
    echo Failed to create the virtual environment.
    pause
    exit /b 1
)

echo Updating pip...
".venv\Scripts\python.exe" -m pip install --upgrade pip

echo Installing application requirements...
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo Installation failed. Check your internet connection and try again.
    pause
    exit /b 1
)

echo.
echo Setup completed successfully.
echo Double-click run_windows.bat to start the software.
pause
