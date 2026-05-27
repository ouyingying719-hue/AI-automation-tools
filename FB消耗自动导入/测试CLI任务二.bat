@echo off
chcp 936 >nul
cd /d "%~dp0"
echo ============================================================
echo 测试 CLI 任务二（fire-and-forget + 等5分钟 + post_verify）
echo 全程约 6 分钟
echo ============================================================
echo.
python -u cli_branch\上传BI_任务二_cli.py
echo.
echo Return code: %errorlevel%
pause
