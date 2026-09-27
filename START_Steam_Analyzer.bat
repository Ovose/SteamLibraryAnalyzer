@echo off
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0SteamLibraryAnalyzer.ps1"
echo.
echo Press any key to close...
pause >nul
