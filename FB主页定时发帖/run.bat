@echo off
chcp 65001 >nul
cd /d "%~dp0"

if not exist config.json (
    echo [ERROR] config.json not found. Copy config.example.json to config.json first.
    pause
    exit /b 1
)

if not exist posts.xlsx (
    echo [ERROR] posts.xlsx not found. Run: python make_template.py
    pause
    exit /b 1
)

echo === DRY RUN ===
python fb_scheduler.py --dry-run
if errorlevel 1 (
    echo.
    echo [ERROR] Dry-run failed. Check the messages above.
    pause
    exit /b 1
)

echo.
set /p CONFIRM=Plan looks good? Press Y to publish, any other key to cancel:
if /i "%CONFIRM%"=="Y" (
    python fb_scheduler.py
) else (
    echo Cancelled.
)
pause
