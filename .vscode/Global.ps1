[string]$Desktop = [Environment]::GetFolderPath('DesktopDirectory')
[string]$WDADesktop = "C:\Users\WDAGUtilityAccount\Desktop"
[string]$Win32App = "C:\M365\win32app"
[string]$Application = "$($env:APPVEYOR_REPO_BRANCH)"
[string]$Cache = "C:\M365\win32app\$Application"
[string]$LogonCommand = "LogonCommand.ps1"

# Cache resources
Remove-Item -Path "$Win32App" -Recurse -Force -ErrorAction Ignore
Copy-Item -Path "Toolkit" -Destination "$Cache" -Recurse -Force -Verbose -ErrorAction Ignore
explorer "$Cache"
