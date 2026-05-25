@echo off
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] python not found in PATH. Please install Python 3.8+ first.
    pause
    exit /b 1
)

python -c "import openpyxl, requests" 2>nul
if errorlevel 1 (
    echo Installing dependencies...
    python -m pip install --quiet openpyxl requests
)

python "%~dp0fb_preview_updater.py" %*
echo.
echo Done. exit=%errorlevel%
pause
