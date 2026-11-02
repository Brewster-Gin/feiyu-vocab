@echo off
chcp 65001 >nul
net session >nul 2>&1
if %errorlevel% neq 0 (
    powershell -NoProfile -Command "try { Start-Process -FilePath '%~f0' -Verb RunAs } catch { }"
    exit /b
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0fix-paths.ps1"
