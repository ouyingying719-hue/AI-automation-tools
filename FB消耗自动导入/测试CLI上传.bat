@echo off
chcp 936 >nul
cd /d "%~dp0"
echo ============================================================
echo CLI Test: Upload via HTTP API (no browser needed)
echo ============================================================
echo.
python -u cli_branch\ÉÏ´«BI_cli.py
echo.
echo Return code: %errorlevel%
pause
