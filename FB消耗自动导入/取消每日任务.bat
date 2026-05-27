@echo off
chcp 936 >nul
echo 删除每日自动任务...
schtasks /Delete /TN "BI每日自动导入" /F
if errorlevel 1 (
    echo [错误] 删除失败
) else (
    echo [OK] 已删除
)
pause
