# Variables
. ".vscode\Global.ps1"

# IntuneWin tool
[string]$Uri = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master"
[string]$Exe = "IntuneWinAppUtil.exe"

$win32appFolderPath = "C:\Temp\win32app"

# Ensure the folder path exists and clean it if necessary
if (Test-Path $win32appFolderPath) {
    Remove-Item -Path $win32appFolderPath -Recurse -Force
}

New-Item -ItemType Directory -Path $win32appFolderPath -Force

# Download the content prep tool if it does not exist
if (-not (Test-Path -Path "$win32appFolderPath\$Exe")) {
    Invoke-WebRequest -Uri "$Uri/$Exe" -OutFile "$win32appFolderPath\$Exe"
}

$IntuneUtilPath = "$win32appFolderPath\$Exe"


# Run the IntuneWinAppUtil tool
Set-Location -Path $win32appFolderPath
& $IntuneUtilPath -c $ToolkitPath -s $InstallexeName -o $win32appFolderPath -q

