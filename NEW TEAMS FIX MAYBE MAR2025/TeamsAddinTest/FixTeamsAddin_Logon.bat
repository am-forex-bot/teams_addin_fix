@echo off
REM ==========================================================================
REM  FixTeamsAddin_Logon.bat
REM  Runs the per-user Teams add-in health-check/fix HIDDEN at logon.
REM  It is self-locating: it runs the .ps1 and .vbs sitting next to it, so the
REM  whole folder can live on a network share or be copied anywhere.
REM  Use this as a GPO/Group Policy logon script if you are NOT using the
REM  scheduled-task installer (Install-TeamsAddinFix.ps1).
REM ==========================================================================
wscript.exe //B //Nologo "%~dp0Run-Hidden.vbs" powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Fix-TeamsMeetingAddin.ps1"
exit /b 0
