@echo off
chcp 936 >nul
cd /d "%~dp0"
echo 手动测试全流程（任务一+任务二）...
echo.
python -u run_all_v2.py
echo.
echo Return code: %errorlevel%
pause
