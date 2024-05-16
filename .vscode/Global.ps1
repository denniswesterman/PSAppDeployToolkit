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

# Variables for sandbox paths
[string]$WDADesktop = "C:\Users\WDAGUtilityAccount\Desktop"
[string]$LogonCommand = "LogonCommand.ps1"
[string]$Win32App = "C:\Temp\win32app"
