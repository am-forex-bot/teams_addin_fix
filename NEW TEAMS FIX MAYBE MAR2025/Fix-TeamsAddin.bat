@echo off
REM ==========================================================================
REM  Fix-TeamsAddin.bat  -  robust Teams Meeting add-in fix (pure batch)
REM
REM  No PowerShell, so it runs under Constrained Language Mode / AppLocker.
REM  Run it from C:\Windows\System32. No admin needed (unless it must install).
REM
REM  It fixes the two things that actually break this add-in:
REM    1. Teams updates to a new version folder -> re-registers WHATEVER
REM       version is present now (so version numbers never matter again).
REM    2. Outlook auto-disables it -> forces it enabled, tells Outlook never
REM       to disable it, and clears any existing disabled/crash entry.
REM
REM  Run with no arguments for a normal (visible) fix.
REM  Run "Fix-TeamsAddin.bat /quiet" for the silent logon self-heal.
REM ==========================================================================
setlocal enableextensions
set "QUIET="
if /i "%~1"=="/quiet" set "QUIET=1"
if not defined QUIET title Teams Meeting Add-in Fix
if not defined QUIET echo Fixing the Teams Meeting add-in for %USERNAME% ...

REM --- Office bitness (default x64) ---
set "PLAT=x64"
for /f "tokens=3" %%P in ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v Platform 2^>nul ^| find /i "Platform"') do set "PLAT=%%P"
if /i not "%PLAT%"=="x86" if /i not "%PLAT%"=="x64" set "PLAT=x64"

REM --- find the newest installed add-in loader DLL ---
call :locate
if defined DLL goto register

REM --- none present: don't try to install at logon (no UAC there) ---
if defined QUIET goto done
if exist "%~dp0MicrosoftTeamsMeetingAddinInstaller.msi" (
    echo No add-in found - installing from the MSI...
    msiexec.exe /i "%~dp0MicrosoftTeamsMeetingAddinInstaller.msi" /qn /norestart
    call :locate
)
if not defined DLL (
    echo Could not find or install the add-in. Make sure new Teams is installed and signed in.
    goto done
)

:register
if not defined QUIET echo Registering: %DLL%
set "REGSVR=%WINDIR%\System32\regsvr32.exe"
if /i "%PLAT%"=="x86" set "REGSVR=%WINDIR%\SysWOW64\regsvr32.exe"
"%REGSVR%" /s /n /i:user "%DLL%"

REM --- enable it for this user ---
reg add "HKCU\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect" /v LoadBehavior /t REG_DWORD /d 3 /f >nul

REM --- stop Outlook ever auto-disabling it, and clear any existing disable ---
reg add "HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList" /v "TeamsAddin.FastConnect" /t REG_DWORD /d 1 /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\CrashingAddinList" /f >nul 2>&1

if not defined QUIET (
    echo.
    echo DONE - add-in registered and enabled for %USERNAME%.
    echo Close Outlook completely and reopen it.
)
goto done

REM ---------- helpers ----------
:locate
set "DLL="
call :scan "%LOCALAPPDATA%\Microsoft\TeamsMeetingAddin"
if defined DLL goto :eof
call :scan "%LOCALAPPDATA%\Microsoft\TeamsMeetingAdd-in"
goto :eof

:scan
if not exist "%~1" goto :eof
for /f "delims=" %%D in ('dir /b /ad /o-d "%~1" 2^>nul') do (
    if not defined DLL if exist "%~1\%%D\%PLAT%\Microsoft.Teams.AddinLoader.dll" set "DLL=%~1\%%D\%PLAT%\Microsoft.Teams.AddinLoader.dll"
    if not defined DLL if exist "%~1\%%D\x64\Microsoft.Teams.AddinLoader.dll" set "DLL=%~1\%%D\x64\Microsoft.Teams.AddinLoader.dll"
    if not defined DLL if exist "%~1\%%D\x86\Microsoft.Teams.AddinLoader.dll" set "DLL=%~1\%%D\x86\Microsoft.Teams.AddinLoader.dll"
)
goto :eof

:done
if not defined QUIET (
    echo.
    pause
)
endlocal
