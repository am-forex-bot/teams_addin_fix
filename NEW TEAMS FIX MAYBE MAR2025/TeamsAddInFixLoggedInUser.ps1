#Store logged in user in $originalUser variable
$originalUser = $env:USERNAME
#Set $outlookpath variable to Outlook.exe for re-opening Outlook at the end
$outlookpath = "C:\Program Files\Microsoft Office\root\Office16\Outlook.exe"


# Output the username to verify it's correct
Write-Host "Logged-in user: $originalUser"

# Function to check if Outlook is running and close it
function Close-Outlook {
    $outlook = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
    if ($outlook) {
        Stop-Process -Name OUTLOOK -Force
        Start-Sleep -Seconds 5  # Wait for 5 seconds to make sure Outlook is closed
        Write-Host "Outlook closed."
    
}
}

# Call the function to close Outlook
Close-Outlook

# Path to the admin script to run admin tasks
$adminScript = "C:\Temp\NEW TEAMS FIX MAYBE MAR2025\TeamsAddInFixAdmin.ps1"

# Run the admin script as administrator, passing the logged-in username as an argument
Write-Host "Running admin script to manage Teams Add-in installation..."
& Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -NoProfile -File `"$adminScript`" -originalUser `"$originalUser`"" -Verb RunAs -Wait
Write-Host "Admin script completed."



# Function to reopen Outlook
function Reopen-Outlook {
    Write-Host "Reopening Outlook as $originalUser..."
    Start-Process -FilePath $outlookpath
    Write-Host "Outlook reopened."
    
}

Reopen-Outlook
