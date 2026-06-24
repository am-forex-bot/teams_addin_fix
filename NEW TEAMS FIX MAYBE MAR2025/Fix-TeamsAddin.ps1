<#
  Fix-TeamsAddin.ps1  -  ONE-STOP Teams Meeting add-in fix

  Designed to be copied into C:\Windows\System32 (alongside Fix-TeamsAddin.bat
  and MicrosoftTeamsMeetingAddinInstaller.msi) on locked-down machines that only
  allow scripts to run from System32. It also runs fine straight from a folder.

  What it does:
    1. Detects the logged-in user automatically (whoever runs it).
    2. If the add-in is already on the PC  -> registers it (no admin needed).
    3. If the add-in is MISSING            -> installs it from the .msi sitting
       next to it, into THIS user's profile (asks for admin only if needed).
    4. Enables it in Outlook (LoadBehavior = 3).
    5. Never hard-codes a version; sets a logon task so it re-checks itself and
       won't break when Teams updates to a new version.

  -Quiet is used by the automatic logon re-check; a normal run is verbose.
#>
[CmdletBinding()]
param([switch]$Quiet)

$ErrorActionPreference = 'Stop'
$Interactive = -not $Quiet

function Info($m){ if($Interactive){ Write-Host $m } }
function Good($m){ if($Interactive){ Write-Host $m -ForegroundColor Green } }
function Warn($m){ if($Interactive){ Write-Host $m -ForegroundColor Yellow } }
function Bad ($m){ if($Interactive){ Write-Host $m -ForegroundColor Red } }

$ProgId  = 'TeamsAddin.FastConnect'
$DllName = 'Microsoft.Teams.AddinLoader.dll'
$MsiName = 'MicrosoftTeamsMeetingAddinInstaller.msi'
$TaskName = 'TeamsMeetingAddinFix'

function Get-Platform {
    $plat = 'x64'
    try {
        $p = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction Stop).Platform
        if ($p -eq 'x86' -or $p -eq 'x64') { $plat = $p }
    } catch { }
    return $plat
}

function Find-LoaderDll {
    $bases = @(
        "$env:LOCALAPPDATA\Microsoft\TeamsMeetingAddin",
        "$env:LOCALAPPDATA\Microsoft\TeamsMeetingAdd-in"
    ) | Where-Object { Test-Path $_ }

    $newest = $bases |
        ForEach-Object { Get-ChildItem $_ -Directory -ErrorAction SilentlyContinue } |
        Where-Object { $_.Name -match '^\d+\.\d+\.\d+\.\d+$' } |
        Sort-Object { [version]$_.Name } -Descending | Select-Object -First 1
    if (-not $newest) { return $null }

    foreach ($p in @((Get-Platform), 'x64', 'x86')) {
        $cand = Join-Path $newest.FullName "$p\$DllName"
        if (Test-Path $cand) {
            return [pscustomobject]@{ Dll = $cand; Version = $newest.Name; Platform = $p }
        }
    }
    return $null
}

function Find-Msi {
    # Non-recursive on purpose: just look next to this script (and one level up).
    # (Never recurse System32.)
    foreach ($dir in @($PSScriptRoot, (Split-Path $PSScriptRoot -Parent))) {
        if ($dir) {
            $cand = Join-Path $dir $MsiName
            if (Test-Path $cand) { return $cand }
        }
    }
    return $null
}

function Get-MsiVersion($msi) {
    try {
        $wi  = New-Object -ComObject WindowsInstaller.Installer
        $db  = $wi.GetType().InvokeMember('OpenDatabase','InvokeMethod',$null,$wi,@($msi,0))
        $vw  = $db.GetType().InvokeMember('OpenView','InvokeMethod',$null,$db,@("SELECT Value FROM Property WHERE Property='ProductVersion'"))
        $vw.GetType().InvokeMember('Execute','InvokeMethod',$null,$vw,$null) | Out-Null
        $rec = $vw.GetType().InvokeMember('Fetch','InvokeMethod',$null,$vw,$null)
        if ($rec) { return $rec.GetType().InvokeMember('StringData','GetProperty',$null,$rec,1) }
    } catch { }
    return $null
}

function Install-FromMsi($msi) {
    $log = Join-Path $env:TEMP 'TeamsAddinMsi.log'

    # First try a per-user install - this usually needs NO admin.
    Info "Installing the add-in (this may take a moment)..."
    Start-Process msiexec.exe -Wait -ArgumentList ('/i "{0}" /qn /norestart MSIINSTALLPERUSER=1 ALLUSERS=2 /l*v "{1}"' -f $msi, $log)
    if (Find-LoaderDll) { return $true }

    # Per-user didn't place the files - retry elevated, but force it into THIS
    # user's profile so the add-in ends up where the logged-in user can see it.
    Warn "Need administrator rights to install - approve the prompt (it installs for $env:USERNAME)..."
    $ver    = Get-MsiVersion $msi
    $target = if ($ver) { "$env:LOCALAPPDATA\Microsoft\TeamsMeetingAddin\$ver" } else { $null }
    $args   = if ($target) {
        '/i "{0}" TARGETDIR="{1}" /qn /norestart /l*v "{2}"' -f $msi, $target, $log
    } else {
        '/i "{0}" /qn /norestart /l*v "{1}"' -f $msi, $log
    }
    try { Start-Process msiexec.exe -ArgumentList $args -Verb RunAs -Wait }
    catch { Bad "The admin prompt was cancelled."; return $false }
    return [bool](Find-LoaderDll)
}

function Register-Loader($info) {
    $regsvr32 = if ($info.Platform -eq 'x64') {
        "$env:WINDIR\System32\regsvr32.exe"
    } else {
        "$env:WINDIR\SysWOW64\regsvr32.exe"
    }
    # /s silent, /n no DllRegisterServer, /i:user = register for THIS user (no admin)
    Start-Process $regsvr32 -Wait -ArgumentList ('/s /n /i:user "{0}"' -f $info.Dll)

    $key = "HKCU:\Software\Microsoft\Office\Outlook\Addins\$ProgId"
    if (-not (Test-Path $key)) { New-Item $key -Force | Out-Null }
    New-ItemProperty $key -Name 'LoadBehavior' -PropertyType DWord -Value 3 -Force | Out-Null
}

function Install-SelfHeal {
    # Register a per-user logon task that re-runs THIS script (-Quiet) from
    # wherever it currently lives. Run from System32 so it stays allowed by
    # policy. Best-effort: if task creation is blocked, the one-off fix still
    # worked. Returns $true only if the task was created.
    try {
        if (-not $PSCommandPath) { return $false }
        $script  = (Resolve-Path $PSCommandPath).Path
        $action  = New-ScheduledTaskAction -Execute 'powershell.exe' `
                    -Argument ('-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "{0}" -Quiet' -f $script)
        $trigger = New-ScheduledTaskTrigger -AtLogOn -User ("{0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
        $set     = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 5)
        Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $set -Force | Out-Null
        return $true
    } catch { return $false }
}

# ===================== main =====================
Info '==================================================='
Info ' Teams Meeting Add-in fix'
Info " User: $env:USERNAME"
Info '==================================================='

$info = Find-LoaderDll
$preexisting = [bool]$info

if (-not $info -and $Interactive) {
    $msi = Find-Msi
    if (-not $msi) {
        Bad "The add-in isn't installed and I can't find $MsiName next to this script."
        Bad "Copy the .msi into the same folder as this script (e.g. System32) and re-run."
        exit 1
    }
    if (Install-FromMsi $msi) { $info = Find-LoaderDll }
}

if (-not $info) {
    if ($Interactive) {
        Bad "Couldn't get the add-in in place. Details in: $env:TEMP\TeamsAddinMsi.log"
    }
    exit 1
}

Register-Loader $info
$healed = $false
if ($Interactive) { $healed = Install-SelfHeal }

$how = if ($preexisting) { 'was already on the PC' } else { 'installed from the MSI' }
Good "SUCCESS - Teams Meeting add-in v$($info.Version) ($($info.Platform)) $how, now registered and enabled."
Info "Registered: $($info.Dll)"
Info ''
Info "Now CLOSE Outlook completely and reopen it - the 'Teams Meeting' button"
Info "will be on the calendar ribbon when you create a new meeting."
if ($healed) {
    Info "(A logon task was set up, so it will re-check itself and won't break when Teams updates.)"
} else {
    Info "(Heads-up: couldn't set the auto re-check task on this PC - just re-run this if the button ever disappears.)"
}
exit 0
