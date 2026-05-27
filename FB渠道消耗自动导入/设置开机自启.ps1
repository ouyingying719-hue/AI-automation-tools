# Create startup shortcut for BI Browser
$startupDir = [Environment]::GetFolderPath('Startup')
$shortcutPath = Join-Path $startupDir 'BI Browser.lnk'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$targetBat = Join-Path $scriptDir '启动主浏览器.bat'

if (-not (Test-Path $targetBat)) {
    Write-Host "[ERROR] Cannot find: $targetBat" -ForegroundColor Red
    exit 1
}

$ws = New-Object -ComObject WScript.Shell
$sc = $ws.CreateShortcut($shortcutPath)
$sc.TargetPath = $targetBat
$sc.WorkingDirectory = $scriptDir
$sc.WindowStyle = 7  # 7 = Minimized
$sc.Save()

if (Test-Path $shortcutPath) {
    Write-Host "[OK] Startup shortcut created" -ForegroundColor Green
    Write-Host "Location: $shortcutPath"
    Write-Host ""
    Write-Host "Next boot, BI Browser will auto-start (minimized)"
} else {
    Write-Host "[FAIL] Could not create shortcut" -ForegroundColor Red
}
