@echo off
chcp 65001 >nul
echo ==== 1. CET6-Word-Reminder details ====
schtasks /query /tn "CET6-Word-Reminder" /v /fo LIST
echo.
echo ==== 2. lines mentioning CET6 (one line per TRIGGER, not per task) ====
schtasks /query /fo LIST | findstr /i "CET6"
echo.
echo ==== 3. real task count, should print 1 ====
schtasks /query /tn "CET6-Word-Reminder" /xml | find /c "<Task "
echo.
echo ==== 4. startup folder (should be empty) ====
dir /b "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
echo.
pause
