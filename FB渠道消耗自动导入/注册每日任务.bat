@echo off
chcp 936 >nul
cd /d "%~dp0"
echo ========================================
echo 注册 Windows 计划任务：每天 08:00 自动运行
echo ========================================
echo.
set TASK_NAME=BI每日自动导入
set BAT_PATH=%~dp0每日自动运行.bat
echo 任务名：%TASK_NAME%
echo 脚本：%BAT_PATH%
echo.
echo 删除已存在的同名任务（如果有）...
schtasks /Delete /TN "%TASK_NAME%" /F >nul 2>&1
echo.
echo 创建新任务...
schtasks /Create /TN "%TASK_NAME%" /TR "\"%BAT_PATH%\"" /SC DAILY /ST 08:00 /RL HIGHEST /F
if errorlevel 1 (
    echo.
    echo [错误] 创建失败，请用管理员权限重试
    pause
    exit /b 1
)
echo.
echo ========================================
echo 注册成功！每天 08:00 自动运行
echo.
echo 查看任务：
schtasks /Query /TN "%TASK_NAME%" /V /FO LIST | findstr /I "任务名 状态 上次 下次"
echo ========================================
pause
