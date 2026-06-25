@echo off
REM ==========================================================================
REM  Fix-TeamsAddin.bat (v2) - Teams Meeting add-in fix, pure batch (CLM-safe)
REM
REM  Synthesised from BOTH earlier fixes, keeping what worked and dropping what
REM  broke:
REM
REM   ORIGINAL (TeamsAddInFixAdmin.ps1):
REM     + uninstalled + REINSTALLED the add-in from the MSI (a real reset)
REM     + set LoadBehavior and bounced Outlook so it actually took effect
REM     - HARD-CODED the version path  ...\TeamsMeetingAddin\1.0.24313.1
REM     - FORCED TARGETDIR to a custom path (breaks the loader's version pick)
REM     - per-machine (HKLM/ALLUSERS=1), needed admin every time
REM
REM   v5 (Fix-TeamsMeetingAddin.ps1):
REM     + dynamic: found the NEWEST version, no hard-coding
REM     + per-user re-register, no admin
REM     - PowerShell (dies under your Constrained Language Mode)
REM     - regsvr32 without /s (popup); loose LoadBehavior pattern matching
REM
REM  This file: dynamic newest-version (no hard-coding), re-registers the
REM  loader for THIS user, sets LoadBehavior in BOTH HKCU and HKLM, clears
REM  Outlook's auto-disable, and (interactively) bounces Outlook so it applies.
REM
REM  Modes:
REM    (no arg)  normal  - re-register + enable + restart Outlook
REM    /repair           - uninstall the add-in and REINSTALL it from the MSI to
REM                        its DEFAULT location (no forced path). Use this for the
REM                        "shows Active in COM add-ins but no button" case - that
REM                        is usually a bad/custom-path install from the old fix.
REM    /quiet            - silent logon self-heal (re-register + enable only)
REM ==========================================================================
setlocal enableextensions
set "MODE=normal"
if /i "%~1"=="/quiet"  set "MODE=quiet"
if /i "%~1"=="/repair" set "MODE=repair"
if not "%MODE%"=="quiet" echo Teams Meeting add-in fix for %USERNAME%  (mode: %MODE%)

REM --- Office bitness, and the matching regsvr32 ---
set "PLAT=x64"
for /f "tokens=3" %%P in ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v Platform 2^>nul ^| find /i "Platform"') do set "PLAT=%%P"
if /i not "%PLAT%"=="x86" if /i not "%PLAT%"=="x64" set "PLAT=x64"
set "REGSVR=%WINDIR%\System32\regsvr32.exe"
if /i "%PLAT%"=="x86" set "REGSVR=%WINDIR%\SysWOW64\regsvr32.exe"

REM --- close Outlook so changes load cleanly (not at logon) ---
if not "%MODE%"=="quiet" (
    echo Closing Outlook so the change can take effect...
    taskkill /im outlook.exe /f >nul 2>&1
)

REM --- /repair: clean reinstall from the MSI to its DEFAULT location ---
if "%MODE%"=="repair" if exist "%~dp0MicrosoftTeamsMeetingAddinInstaller.msi" (
    echo Removing the current add-in and reinstalling it cleanly...
    msiexec.exe /x "%~dp0MicrosoftTeamsMeetingAddinInstaller.msi" /qn /norestart
    msiexec.exe /i "%~dp0MicrosoftTeamsMeetingAddinInstaller.msi" /qn /norestart
)

REM --- find newest installed loader DLL (any version, both folder names) ---
call :locate
if not defined DLL if not "%MODE%"=="quiet" if exist "%~dp0MicrosoftTeamsMeetingAddinInstaller.msi" (
    echo No add-in present - installing from MSI...
    msiexec.exe /i "%~dp0MicrosoftTeamsMeetingAddinInstaller.msi" /qn /norestart
    call :locate
)
if not defined DLL (
    if not "%MODE%"=="quiet" echo Could not find or install the add-in.
    goto end
)

REM --- re-register the loader for THIS user (this is what keeps it Active) ---
if not "%MODE%"=="quiet" echo Registering: %DLL%
"%REGSVR%" /s /n /i:user "%DLL%"

REM --- enable it, both hives (original used HKLM, v5 used HKCU - do both) ---
reg add "HKCU\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect" /v LoadBehavior /t REG_DWORD /d 3 /f >nul 2>&1
reg add "HKLM\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect" /v LoadBehavior /t REG_DWORD /d 3 /f >nul 2>&1

REM --- stop Outlook auto-disabling it, and clear any existing disable ---
reg add "HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList" /v "TeamsAddin.FastConnect" /t REG_DWORD /d 1 /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\CrashingAddinList" /f >nul 2>&1

REM --- reopen Outlook so it reloads the add-in (not at logon) ---
if not "%MODE%"=="quiet" start "" outlook.exe

if not "%MODE%"=="quiet" (
    echo.
    echo DONE - registered %DLL%
    echo Outlook has been restarted - check Calendar ^> New Meeting for the Teams button.
)
goto end

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

:end
if not "%MODE%"=="quiet" (
    echo.
    pause
)
endlocal
