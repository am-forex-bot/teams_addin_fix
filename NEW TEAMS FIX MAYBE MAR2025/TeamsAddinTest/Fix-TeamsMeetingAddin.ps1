<#
  Fix-TeamsMeetingAddin.ps1  (v6)

  Purpose
  -------
  Makes the "Teams Meeting" add-in show up in Outlook again, for the CURRENT
  user, with NO admin rights and NO UAC prompt, using whatever version of the
  add-in is currently installed on the machine.

  How it works
  ------------
  1. Finds the newest installed add-in version under the user's LocalAppData
     (handles both "TeamsMeetingAddin" and "TeamsMeetingAdd-in" folder names).
  2. Health-check FIRST: if the add-in is already correctly registered to that
     exact version AND LoadBehavior = 3, it does NOTHING and exits fast. This is
     what keeps logon quick - it only acts when something is actually wrong.
  3. If a fix is needed, it silently re-registers the loader DLL per-user
     (regsvr32 /s /n /i:user), forces LoadBehavior = 3, and tells Outlook not to
     auto-disable the add-in.

  It NEVER closes Outlook and ALWAYS exits 0, so it can run at logon without
  interrupting or blocking the user. Changes appear the next time Outlook starts.

  Parameters
  ----------
  -Force        Re-register even if it looks healthy, and also clear Outlook's
                disabled/crashing add-in lists (use for manual "fix it now" runs).
  -ShowConsole  Print what it's doing to the console (for manual troubleshooting).

  Logs to:  %LOCALAPPDATA%\Company\TeamsAddinFix.log     (rolling, fixes/errors)
            %LOCALAPPDATA%\Company\TeamsAddinFix.status   (last run heartbeat)
#>
[CmdletBinding()]
param(
    [switch]$Force,
    [switch]$ShowConsole
)

$ErrorActionPreference = 'Stop'

# ---- constants ----
$ProgId        = 'TeamsAddin.FastConnect'
$LoaderDllName = 'Microsoft.Teams.AddinLoader.dll'
$AddinKey      = "HKCU:\Software\Microsoft\Office\Outlook\Addins\$ProgId"

# ---- logging / output helpers ----
$LogDir     = Join-Path $env:LOCALAPPDATA 'Company'
$LogFile    = Join-Path $LogDir 'TeamsAddinFix.log'
$StatusFile = Join-Path $LogDir 'TeamsAddinFix.status'

function Ensure-LogDir {
    if (-not (Test-Path -LiteralPath $LogDir)) {
        New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
    }
}
function Write-Log([string]$msg) {
    try {
        Ensure-LogDir
        $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Add-Content -LiteralPath $LogFile -Value "$ts  $msg" -Encoding UTF8
        $lines = @(Get-Content -LiteralPath $LogFile -ErrorAction SilentlyContinue)
        if ($lines.Count -gt 200) { $lines[-200..-1] | Set-Content -LiteralPath $LogFile -Encoding UTF8 }
    } catch { }
}
function Write-Status([string]$state, [string]$detail) {
    try {
        Ensure-LogDir
        $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        "$ts  $state  $detail" | Set-Content -LiteralPath $StatusFile -Encoding UTF8
    } catch { }
}
function Say([string]$msg) {
    if ($ShowConsole) { Write-Host $msg }
}

# ---- registry read helpers ----
function Get-RegDefault([string]$path) {
    try { return (Get-Item -LiteralPath $path -ErrorAction Stop).GetValue('') }
    catch { return $null }
}
function Get-RegValue([string]$path, [string]$name) {
    try { return (Get-ItemProperty -LiteralPath $path -Name $name -ErrorAction Stop).$name }
    catch { return $null }
}

# ---- find newest installed add-in version + its loader DLL ----
function Get-LatestLoaderDll {
    $bases = @(
        (Join-Path $env:LOCALAPPDATA 'Microsoft\TeamsMeetingAddin'),
        (Join-Path $env:LOCALAPPDATA 'Microsoft\TeamsMeetingAdd-in')
    ) | Where-Object { Test-Path -LiteralPath $_ }
    if (-not $bases) { return $null }

    $versions = foreach ($base in $bases) {
        Get-ChildItem -LiteralPath $base -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^\d+\.\d+\.\d+\.\d+$' } |
            ForEach-Object {
                [PSCustomObject]@{ Version = [version]$_.Name; FullPath = $_.FullName }
            }
    }
    if (-not $versions) { return $null }

    # Office bitness (Click-to-Run); default to x64.
    $platform = 'x64'
    try {
        $p = (Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction Stop).Platform
        if ($p -in @('x86','x64')) { $platform = $p }
    } catch { }

    foreach ($v in ($versions | Sort-Object Version -Descending)) {
        # Prefer the bitness that matches Office, but fall back to the other if needed.
        foreach ($plat in @($platform, $(if ($platform -eq 'x64') { 'x86' } else { 'x64' }))) {
            $dll = Join-Path $v.FullPath (Join-Path $plat $LoaderDllName)
            if (Test-Path -LiteralPath $dll) {
                return [PSCustomObject]@{
                    Dll = $dll; Version = $v.Version.ToString(); Platform = $plat; Path = $v.FullPath
                }
            }
        }
    }
    return $null
}

# ---- which DLL is currently registered for the add-in COM class? ----
function Get-RegisteredLoaderPath {
    $progIdKeys = @(
        "HKCU:\Software\Classes\$ProgId\CLSID",
        "HKCU:\Software\Classes\WOW6432Node\$ProgId\CLSID"
    )
    foreach ($pk in $progIdKeys) {
        $clsid = Get-RegDefault $pk
        if ($clsid) {
            $inprocKeys = @(
                "HKCU:\Software\Classes\CLSID\$clsid\InprocServer32",
                "HKCU:\Software\Classes\WOW6432Node\CLSID\$clsid\InprocServer32"
            )
            foreach ($ik in $inprocKeys) {
                $dll = Get-RegDefault $ik
                if ($dll) { return ([string]$dll).Trim().Trim('"') }
            }
        }
    }
    return $null
}

# ---- fix actions ----
function Register-Loader([string]$dll, [string]$platform) {
    $regsvr32 = if ($platform -eq 'x64') {
        Join-Path $env:WINDIR 'System32\regsvr32.exe'
    } else {
        Join-Path $env:WINDIR 'SysWOW64\regsvr32.exe'
    }
    # /s = silent (no popup), /n = don't call DllRegisterServer, /i:user = per-user install
    $argLine = '/s /n /i:user "{0}"' -f $dll
    $proc = Start-Process -FilePath $regsvr32 -ArgumentList $argLine -Wait -PassThru -WindowStyle Hidden
    return $proc.ExitCode
}
function Set-LoadBehavior {
    if (-not (Test-Path -LiteralPath $AddinKey)) { New-Item -Path $AddinKey -Force | Out-Null }
    New-ItemProperty -Path $AddinKey -Name 'LoadBehavior' -PropertyType DWord -Value 3 -Force | Out-Null
}
function Protect-FromAutoDisable([bool]$clearDisabled) {
    foreach ($ver in @('16.0','15.0')) {
        $outlookKey = "HKCU:\Software\Microsoft\Office\$ver\Outlook"
        if (-not (Test-Path -LiteralPath $outlookKey)) { continue }
        $res = "$outlookKey\Resiliency"
        $dnd = "$res\DoNotDisableAddinList"
        if (-not (Test-Path -LiteralPath $dnd)) { New-Item -Path $dnd -Force | Out-Null }
        New-ItemProperty -Path $dnd -Name $ProgId -PropertyType DWord -Value 1 -Force | Out-Null
        if ($clearDisabled) {
            foreach ($sub in @('DisabledItems','CrashingAddinList')) {
                $p = "$res\$sub"
                if (Test-Path -LiteralPath $p) {
                    Remove-Item -LiteralPath $p -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }
}

# ================= main =================
try {
    Say 'Teams Meeting Add-in fix - checking...'

    $latest = Get-LatestLoaderDll
    if (-not $latest) {
        Write-Status 'NO-ADDIN' 'No TeamsMeetingAddin version/loader DLL found in LocalAppData.'
        Write-Log    'No TeamsMeetingAddin loader DLL found (is new Teams / the add-in installed for this user?). No action.'
        Say          'Could not find the Teams Meeting Add-in files. Is new Teams installed for this user?'
        exit 0
    }

    Say ('Found add-in v{0} ({1}): {2}' -f $latest.Version, $latest.Platform, $latest.Dll)

    $registered = Get-RegisteredLoaderPath
    $lb         = Get-RegValue $AddinKey 'LoadBehavior'

    $regOk = ($registered) -and (Test-Path -LiteralPath $registered) -and ($registered -ieq $latest.Dll)
    $lbOk  = ($lb -eq 3)

    if ($regOk -and $lbOk -and -not $Force) {
        Write-Status 'OK' ('Healthy v{0}; no action.' -f $latest.Version)
        Say 'All good - the add-in is correctly registered and enabled. No changes needed.'
        exit 0
    }

    Say 'Applying fix (re-registering the add-in for this user)...'
    $rc = Register-Loader -dll $latest.Dll -platform $latest.Platform
    Set-LoadBehavior
    Protect-FromAutoDisable -clearDisabled:$Force

    $reason = @()
    if (-not $regOk) { $reason += $(if ($registered) { "stale/missing registration ($registered)" } else { 'not registered' }) }
    if (-not $lbOk)  { $reason += "LoadBehavior=$lb" }
    if ($Force)      { $reason += 'forced' }

    Write-Status 'FIXED' ('v{0} {1}; regsvr32 rc={2}; reason: {3}' -f $latest.Version, $latest.Platform, $rc, ($reason -join ', '))
    Write-Log    ("Re-registered add-in v{0} ({1}) at '{2}'; regsvr32 rc={3}; reason: {4}" -f $latest.Version, $latest.Platform, $latest.Dll, $rc, ($reason -join ', '))
    Say 'Fix applied. If Outlook is open, close and reopen it for the Teams Meeting button to appear.'
    exit 0
}
catch {
    Write-Status 'ERROR' $_.Exception.Message
    Write-Log    ('ERROR: ' + $_.Exception.Message)
    Say          ('Error: ' + $_.Exception.Message)
    exit 0
}
