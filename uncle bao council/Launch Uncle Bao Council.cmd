@echo off
setlocal
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0UncleBaoCouncil.ps1" -Action Launch -HideConsole
