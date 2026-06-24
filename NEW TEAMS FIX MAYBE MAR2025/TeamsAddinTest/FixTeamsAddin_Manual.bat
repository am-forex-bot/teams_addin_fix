@echo off
REM ==========================================================================
REM  FixTeamsAddin_Manual.bat   --   "Fix my Teams add-in now"
REM  Double-click this to force a fix on the machine you are sitting at.
REM  Runs VISIBLY so you can see what happened, forces a re-register, and
REM  clears Outlook's disabled-add-in list. No admin rights required.
REM ==========================================================================
echo Fixing the Teams Meeting add-in for the current user...
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Fix-TeamsMeetingAddin.ps1" -Force -ShowConsole
echo.
echo ------------------------------------------------------------------
echo If Outlook is open, CLOSE and REOPEN it for the Teams Meeting
echo button to appear on the calendar ribbon.
echo ------------------------------------------------------------------
echo.
pause
