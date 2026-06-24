<#
  Fix-TeamsAddin.ps1  -  ONE-STOP Teams Meeting add-in fix

  Just run Fix-TeamsAddin.bat. This script:
    1. Detects the logged-in user automatically (whoever runs it).
    2. If the add-in is already on the PC  -> registers it (no admin needed).
    3. If the add-in is MISSING            -> installs it from
       MicrosoftTeamsMeetingAddinInstaller.msi, into THIS user's profile.
       It only asks for admin if the install actually needs it.
    4. Enables it in Outlook (LoadBehavior = 3).
    5. Works with any version (never hard-codes a version number) and sets
       itself to re-check at each logon, so it won't break when Teams updates.

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
    foreach ($root in @($PSScriptRoot, (Split-Path $PSScriptRoot -Parent))) {
        if ($root) {
            $hit = Get-ChildItem -Path $root -Recurse -Filter $MsiName -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($hit) { return $hit.FullName }
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
    try {
        $dir = Join-Path $env:LOCALAPPDATA 'TeamsAddinFix'
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        $dest = Join-Path $dir 'Fix-TeamsAddin.ps1'
        if ($PSCommandPath -and ((Resolve-Path $PSCommandPath).Path -ne $dest)) {
            Copy-Item -LiteralPath $PSCommandPath -Destination $dest -Force
        }
        $startup = [Environment]::GetFolderPath('Startup')
        $vbs     = Join-Path $startup 'TeamsAddinFix.vbs'
        $launch  = 'CreateObject("WScript.Shell").Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""' + $dest + '"" -Quiet", 0, False'
        Set-Content -LiteralPath $vbs -Value $launch -Encoding ASCII
    } catch { }
}

# ===================== main =====================
Info '==================================================='
Info ' Teams Meeting Add-in fix'
Info " User: $env:USERNAME"
Info '==================================================='

$info = Find-LoaderDll

if (-not $info -and $Interactive) {
    $msi = Find-Msi
    if (-not $msi) {
        Bad "The add-in isn't installed and I can't find $MsiName next to this script."
        Bad "Make sure the whole folder (including the .msi) was copied to this PC."
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
if ($Interactive) { Install-SelfHeal }

Good "SUCCESS - Teams Meeting add-in v$($info.Version) is installed, registered and enabled."
Info ''
Info "Now CLOSE Outlook completely and reopen it - the 'Teams Meeting' button"
Info "will be on the calendar ribbon when you create a new meeting."
Info "(It will also re-check itself at each logon, so it won't break when Teams updates.)"
exit 0
