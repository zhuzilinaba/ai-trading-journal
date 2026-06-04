@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo Starting AI Trading Journal...
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [Error] Python was not found. Please install Python 3.10 or newer.
    echo Download: https://www.python.org/downloads/
    pause
    exit /b 1
)

if not exist ".venv" (
    echo Creating virtual environment...
    python -m venv .venv
)

call .venv\Scripts\activate.bat

echo Installing dependencies...
python -m pip install --upgrade pip -q
python -m pip install -r requirements.txt -q

python trade_journal.py

pause
