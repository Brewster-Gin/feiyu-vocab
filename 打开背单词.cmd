@echo off
rem Open the word popup right now (ignores "already done today"), no console flash.
start "" wscript.exe "%~dp0launch.vbs" preview
