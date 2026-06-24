@echo off
title Teams Meeting Add-in Fix
echo Fixing the Teams Meeting add-in for %USERNAME%...
echo (If Outlook is open, please close it first.)
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Fix-TeamsAddin.ps1"
echo.
pause
