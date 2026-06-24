param (
    [string]$originalUser  # Receive the logged-in user from the main script
)

# Path to Teams Add-in Installer
$msiPath = "C:\Temp\NEW TEAMS FIX MAYBE MAR2025\MicrosoftTeamsMeetingAddinInstaller.msi"
$addinName = "Microsoft Teams Meeting Add-in for Microsoft Office"
$addInInstallDir = "C:\Users\$originalUser\AppData\Local\Microsoft\TeamsMeetingAddin\1.0.24313.1"
$addinRegistryKey = "HKLM:\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect"


# Get the uninstall string for the Teams Add-in from Apps & Features
$addin = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -eq $addinName }

if ($addin) {
    Write-Host "Uninstalling $addinName..."
    Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $($addin.IdentifyingNumber) /qn /norestart" -Wait
    Write-Host "$addinName uninstalled."
} else {
    Write-Host "$addinName not found."
}


Write-Host "Installing Microsoft Teams Meeting Add-in for $originalUser..."

Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiPath`" TARGETDIR=`"$addInInstallDir`" /qn ALLUSERS=1 /norestart" -Wait

Write-Host "Microsoft Teams Meeting Add-in installed."

Start-Sleep -Seconds 5

Set-ItemProperty -Path $addinRegistryKey -Name "LoadBehavior" -Value 3  # 3 means load and enabled

Start-Sleep -Seconds 5
