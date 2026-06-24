@echo off
wscript.exe //B //NoLogo "%SystemRoot%\System32\RunHidden.vbs" powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SystemRoot%\System32\Fix-TeamsMeetingAddin.ps1"
exit /b 0
