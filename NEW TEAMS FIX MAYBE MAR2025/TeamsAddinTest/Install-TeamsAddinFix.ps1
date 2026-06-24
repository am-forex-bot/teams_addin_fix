#Requires -RunAsAdministrator
<#
  Install-TeamsAddinFix.ps1

  Deploys Fix-TeamsMeetingAddin.ps1 so it runs automatically, HIDDEN, in each
  user's own (non-admin) context every time they log on - via a single
  machine-wide Scheduled Task.

  Run this ONCE per machine, as Administrator (manually, or pushed by
  Intune / SCCM / GPO / your RMM tool). After that, every user who logs on to
  that machine gets the quick health-check/fix with no further action.

  Re-run it any time to update the deployed script to a newer version.
#>
param(
    [string]$InstallDir = (Join-Path $env:ProgramData 'TeamsAddinFix'),
    [string]$TaskName   = 'TeamsMeetingAddinFix'
)
$ErrorActionPreference = 'Stop'

$src = Join-Path $PSScriptRoot 'Fix-TeamsMeetingAddin.ps1'
if (-not (Test-Path -LiteralPath $src)) {
    throw "Cannot find Fix-TeamsMeetingAddin.ps1 next to this installer ($PSScriptRoot)."
}

# 1) Copy the fix script to a stable, all-users-readable location.
New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
$ps1 = Join-Path $InstallDir 'Fix-TeamsMeetingAddin.ps1'
Copy-Item -LiteralPath $src -Destination $ps1 -Force

# 2) Create / refresh the scheduled task.
$action  = New-ScheduledTaskAction -Execute 'powershell.exe' `
            -Argument ('-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "{0}"' -f $ps1)

$trigger = New-ScheduledTaskTrigger -AtLogOn   # no -User => fires for ANY user that logs on

# Run as the logged-on user, NOT elevated. S-1-5-32-545 = the built-in "Users"
# group (locale-independent). This is what keeps it per-user and admin-free.
$principal = New-ScheduledTaskPrincipal -GroupId 'S-1-5-32-545' -RunLevel Limited

$settings  = New-ScheduledTaskSettingsSet -StartWhenAvailable `
            -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
            -ExecutionTimeLimit (New-TimeSpan -Minutes 5) -Hidden

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
    -Principal $principal -Settings $settings -Force | Out-Null

Write-Host "Done."
Write-Host "  Scheduled task : $TaskName  (runs hidden at every user logon, non-admin)"
Write-Host "  Fix script     : $ps1"
Write-Host ""
Write-Host "To test now without logging off, run (as the test user):"
Write-Host "  Start-ScheduledTask -TaskName '$TaskName'"
