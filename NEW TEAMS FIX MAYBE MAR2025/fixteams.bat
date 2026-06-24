@echo off
set SCRIPT="C:\Temp\NEW TEAMS FIX MAYBE MAR2025\TeamsAddInFixLoggedInUser.ps1"

where pwsh >nul 2>nul
if %ERRORLEVEL%==0 (
  pwsh -NoProfile -ExecutionPolicy Bypass -File %SCRIPT%
  set EXITCODE=%ERRORLEVEL%
) else (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File %SCRIPT%
  set EXITCODE=%ERRORLEVEL%
)

echo.
echo PowerShell exited with code %EXITCODE%.
echo Press any key to close this window...
pause >nul
