# Vars
# . ".vscode\Global.ps1"

# intunewin
[string]$Uri = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/blob/master"
[string]$Exe = "IntuneWinAppUtil.exe"


$win32appFolderPath = "C:\Temp\win32app"

if (Test-Path $win32appFolderPath) {
    Remove-Item -Path $win32appFolderPath -Recurse -Force
}

New-Item -ItemType Directory -Path $win32appFolderPath -Force

# Source content prep tool
if (-not(Test-Path -Path "$win32appFolderPath\$Exe")){
    Invoke-WebRequest -Uri "$Uri/$Exe" -OutFile "$win32appFolderPath\$Exe"
}
$currentDirectory = Get-Location
# Execute content prep tool
$processOptions = @{
    FilePath = "$win32appFolderPath\$Exe"
    ArgumentList  = "-c ""$Cache"" -s ""$Cache\Deploy-Application.exe"" -o ""$win32appFolderPath"" -q"
    WindowStyle = "Maximized"
    Wait = $true
}

Start-Process @processOptions

# Rename and prepare for upload
Move-Item -Path "$env:TEMP\Deploy-Application.intunewin" -Destination "$Desktop\$Application.intunewin" -Force -Verbose
explorer $Desktop