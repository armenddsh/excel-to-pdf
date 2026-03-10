@echo off
cd /d "%~dp0"

echo Checking dependencies with UV...
if not exist "uv.lock" (
    echo Installing dependencies with UV...
    uv sync
    if errorlevel 1 (
        echo Failed to install dependencies. Please run:
        echo uv sync
        pause
        exit /b 1
    )
)

echo Starting GUI with UV...
uv run gui.py
pause
