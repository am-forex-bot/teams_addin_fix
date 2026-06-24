#Requires -RunAsAdministrator
<#
  Uninstall-TeamsAddinFix.ps1
  Removes the scheduled task and deployed files created by Install-TeamsAddinFix.ps1.
  This does NOT touch the Teams add-in itself - it only stops the auto-fix from running.
#>
param(
    [string]$InstallDir = (Join-Path $env:ProgramData 'TeamsAddinFix'),
    [string]$TaskName   = 'TeamsMeetingAddinFix'
)
$ErrorActionPreference = 'SilentlyContinue'

Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
Remove-Item -LiteralPath $InstallDir -Recurse -Force

Write-Host "Removed scheduled task '$TaskName' and folder '$InstallDir' (if they existed)."
