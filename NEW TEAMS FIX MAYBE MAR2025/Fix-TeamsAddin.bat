@echo off
REM ==========================================================================
REM  Fix-TeamsAddin.bat  -  ONE-STOP Teams Meeting add-in fix (pure batch)
REM
REM  No PowerShell - works on machines locked to Constrained Language Mode.
REM  Uses only built-in System32 tools (reg, regsvr32, msiexec).
REM  Run it from C:\Windows\System32 (where your policy allows .bat to run).
REM
REM  What it does, for the logged-in user, with no admin needed:
REM    - finds the NEWEST installed Teams add-in version (any version)
REM    - if none is installed, installs it from the .msi next to this file
REM    - registers it for this user (regsvr32 /i:user) and enables it in Outlook
REM ==========================================================================
setlocal enableextensions
title Teams Meeting Add-in Fix

echo ===================================================
echo  Teams Meeting Add-in fix
echo  User: %USERNAME%
echo ===================================================
echo.

REM --- detect Office bitness (default x64) ---
set "PLAT=x64"
for /f "tokens=3" %%P in ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v Platform 2^>nul ^| find /i "Platform"') do set "PLAT=%%P"
if /i not "%PLAT%"=="x86" if /i not "%PLAT%"=="x64" set "PLAT=x64"

REM --- look for an already-installed add-in ---
call :locate
if defined DLL set "SRC=was already on the PC"
if defined DLL goto register

REM --- not there: install from the MSI sitting next to this script ---
if not exist "%~dp0MicrosoftTeamsMeetingAddinInstaller.msi" (
    echo The add-in isn't installed yet, and the installer MSI isn't next to this file.
    echo Open Teams and sign in once ^(that installs the add-in^), then run this again.
    goto done
)
echo Add-in not found - installing it from the MSI...
msiexec.exe /i "%~dp0MicrosoftTeamsMeetingAddinInstaller.msi" /qn /norestart
call :locate
if defined DLL set "SRC=installed from the MSI"
if defined DLL goto register

echo Could not install the add-in automatically - it may need an administrator.
echo Try double-clicking MicrosoftTeamsMeetingAddinInstaller.msi to install it,
echo then run this fix again.
goto done

:register
echo Found: %DLL%
set "REGSVR=%WINDIR%\System32\regsvr32.exe"
if /i "%PLAT%"=="x86" set "REGSVR=%WINDIR%\SysWOW64\regsvr32.exe"
"%REGSVR%" /s /n /i:user "%DLL%"
reg add "HKCU\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect" /v LoadBehavior /t REG_DWORD /d 3 /f >nul
echo.
echo SUCCESS - Teams Meeting add-in registered and enabled for %USERNAME% (%SRC%).
echo.
echo Now CLOSE Outlook completely and reopen it - the 'Teams Meeting' button
echo will be on the calendar ribbon when you create a new meeting.
goto done

REM ---------- helpers ----------
:locate
set "DLL="
call :scan "%LOCALAPPDATA%\Microsoft\TeamsMeetingAddin"
if defined DLL goto :eof
call :scan "%LOCALAPPDATA%\Microsoft\TeamsMeetingAdd-in"
goto :eof

:scan
REM %1 = base folder. Walk subfolders newest-first; first one that actually
REM contains the loader DLL wins. ("if not defined DLL" keeps only the first.)
if not exist "%~1" goto :eof
for /f "delims=" %%D in ('dir /b /ad /o-d "%~1" 2^>nul') do (
    if not defined DLL if exist "%~1\%%D\%PLAT%\Microsoft.Teams.AddinLoader.dll" set "DLL=%~1\%%D\%PLAT%\Microsoft.Teams.AddinLoader.dll"
    if not defined DLL if exist "%~1\%%D\x64\Microsoft.Teams.AddinLoader.dll" set "DLL=%~1\%%D\x64\Microsoft.Teams.AddinLoader.dll"
    if not defined DLL if exist "%~1\%%D\x86\Microsoft.Teams.AddinLoader.dll" set "DLL=%~1\%%D\x86\Microsoft.Teams.AddinLoader.dll"
)
goto :eof

:done
echo.
pause
endlocal
