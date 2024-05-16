# Get the current directory
$currentDirectory = Get-Location
Write-Output "Current Directory: $currentDirectory"

# Set the correct path to the Toolkit directory
$ToolkitPath = "$currentDirectory\Toolkit"
if (-not (Test-Path -Path $ToolkitPath)) {
    Write-Error "The path $ToolkitPath does not exist."
    exit
}

# Set the source path and installation script based on the current directory
$AppFullName = Get-Item -Path $ToolkitPath
$Installexe = Get-ChildItem -Path $AppFullName.FullName -Filter *.exe
$InstallexeName = $Installexe.Name

<#
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
#>