@echo off
title Trade Diary — First Time Setup
color 0A
echo.
echo  ============================================================
echo   Trade Diary — Windows App Setup
echo   This will install dependencies and launch the app.
echo  ============================================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python not found!
    echo.
    echo  Please install Python from: https://www.python.org/downloads/
    echo  IMPORTANT: Check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

echo  [1/3] Python found.
echo  [2/3] Installing required packages (first time only)...
pip install pyqt6 pandas openpyxl --quiet --exists-action i

echo  [3/3] Launching Trade Diary...
echo.
python "%~dp0trade_diary.py"

if errorlevel 1 (
    echo.
    echo  [ERROR] App failed to start. Check the error above.
    pause
)
