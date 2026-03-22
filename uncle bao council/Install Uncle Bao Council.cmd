@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0UncleBaoCouncil.ps1" -Action Install
echo.
pause
