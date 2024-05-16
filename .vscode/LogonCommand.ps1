# Determine the path of the global script relative to this script
$globalScriptPath = Join-Path -Path (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent) -ChildPath 'Global.ps1'
if (-not (Test-Path -Path $globalScriptPath)) {
    Write-Error "The global script $globalScriptPath does not exist."
    exit
}
. $globalScriptPath

# Begin the installation process
Start-Process -FilePath "$ToolkitPath\$InstallexeName" -WindowStyle Maximized -Wait



<#
# Check if the installation executable requires winget
if (Get-Content -Path "$ToolkitPath\$InstallexeName" | Select-String "winget install") {
    Write-Host "Requires winget"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $WebClient = New-Object System.Net.WebClient

    # Define resources to download
    $resources = @(
        @{
            fileName = 'Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle'
            url      = 'https://api.github.com/repos/microsoft/winget-cli/releases/latest'
            hash     = 'some_hash_value'  # Placeholder for the actual hash
        },
        @{
            fileName = 'Microsoft.VCLibs.x64.14.00.Desktop.appx'
            url      = 'https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx'
            hash     = 'A39CEC0E70BE9E3E48801B871C034872F1D7E5E8EEBE986198C019CF2C271040'
        },
        @{
            fileName = 'Microsoft.UI.Xaml.2.7.zip'
            url      = 'https://www.nuget.org/api/v2/package/Microsoft.UI.Xaml/2.7.0/'
            hash     = '422FD24B231E87A842C4DAEABC6A335112E0D35B86FAC91F5CE7CF327E36A591'
        }
    )

    # Download and validate resources
    foreach ($resource in $resources) {
        $filePath = Join-Path -Path $env:TEMP -ChildPath $resource.fileName
        try {
            $WebClient.DownloadFile($resource.url, $filePath)
        } catch {
            throw [System.Net.WebException]::new("Download error $($resource.url).", $_.Exception)
        }
        if (-not ($resource.hash -eq (Get-FileHash $filePath).Hash)) {
            throw [System.Activities.VersionMismatchException]::new('Hash mismatch')
        }

        # Extract zip files if necessary
        if ($filePath -match '\.zip$') {
            Expand-Archive -Path $filePath -DestinationPath ($env:TEMP + "\Extracted") -Force
        }
    }

    # Install downloaded packages
    $msixPath = (Join-Path -Path $env:TEMP -ChildPath 'Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle')
    $vcLibsPath = (Join-Path -Path $env:TEMP -ChildPath 'Microsoft.VCLibs.x64.14.00.Desktop.appx')
    $uiLibsPath = (Join-Path -Path $env:TEMP -ChildPath 'Extracted\tools\AppX\x64\Release\Microsoft.UI.Xaml.2.7.appx')

    Add-AppxPackage -Path $msixPath -DependencyPath $vcLibsPath, $uiLibsPath -Verbose
}
#>
