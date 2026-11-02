@echo off
chcp 65001 >nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%LOCALAPPDATA%\CET6Plan\uninstall-package.ps1"
echo.
pause
