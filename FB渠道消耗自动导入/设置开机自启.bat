@echo off
title Setup Auto-Start
cd /d "%~dp0"

echo ========================================
echo   Setup BI Browser Auto-Start
echo ========================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0设置开机自启.ps1"

echo.
pause
