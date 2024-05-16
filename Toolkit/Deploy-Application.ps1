<#
.SYNOPSIS

PSApppDeployToolkit - This script performs the installation or uninstallation of an application(s).

.DESCRIPTION

- The script is provided as a template to perform an install or uninstall of an application(s).
- The script either performs an "Install" deployment type or an "Uninstall" deployment type.
- The install deployment type is broken down into 3 main sections/phases: Pre-Install, Install, and Post-Install.

The script dot-sources the AppDeployToolkitMain.ps1 script which contains the logic and functions required to install or uninstall an application.

PSApppDeployToolkit is licensed under the GNU LGPLv3 License - (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham and Muhammad Mashwani).

This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the
Free Software Foundation, either version 3 of the License, or any later version. This program is distributed in the hope that it will be useful, but
WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License
for more details. You should have received a copy of the GNU Lesser General Public License along with this program. If not, see <http://www.gnu.org/licenses/>.

.PARAMETER DeploymentType

The type of deployment to perform. Default is: Install.

.PARAMETER DeployMode

Specifies whether the installation should be run in Interactive, Silent, or NonInteractive mode. Default is: Interactive. Options: Interactive = Shows dialogs, Silent = No dialogs, NonInteractive = Very silent, i.e. no blocking apps. NonInteractive mode is automatically set if it is detected that the process is not user interactive.

.PARAMETER AllowRebootPassThru

Allows the 3010 return code (requires restart) to be passed back to the parent process (e.g. SCCM) if detected from an installation. If 3010 is passed back to SCCM, a reboot prompt will be triggered.

.PARAMETER TerminalServerMode

Changes to "user install mode" and back to "user execute mode" for installing/uninstalling applications for Remote Desktop Session Hosts/Citrix servers.

.PARAMETER DisableLogging

Disables logging to file for the script. Default is: $false.

.EXAMPLE

powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeployMode 'Silent'; Exit $LastExitCode }"

.EXAMPLE

powershell.exe -Command "& { & '.\Deploy-Application.ps1' -AllowRebootPassThru; Exit $LastExitCode }"

.EXAMPLE

powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeploymentType 'Uninstall'; Exit $LastExitCode }"

.EXAMPLE

Deploy-Application.exe -DeploymentType "Install" -DeployMode "Silent"

.INPUTS

None

You cannot pipe objects to this script.

.OUTPUTS

None

This script does not generate any output.

.NOTES

Toolkit Exit Code Ranges:
- 60000 - 68999: Reserved for built-in exit codes in Deploy-Application.ps1, Deploy-Application.exe, and AppDeployToolkitMain.ps1
- 69000 - 69999: Recommended for user customized exit codes in Deploy-Application.ps1
- 70000 - 79999: Recommended for user customized exit codes in AppDeployToolkitExtensions.ps1

.LINK

https://psappdeploytoolkit.com
#>


[CmdletBinding()]
Param (
    [Parameter(Mandatory = $false)]
    [ValidateSet('Install', 'Uninstall', 'Repair')]
    [String]$DeploymentType = 'Install',
    [Parameter(Mandatory = $false)]
    [ValidateSet('Interactive', 'Silent', 'NonInteractive')]
    [String]$DeployMode = 'Interactive',
    [Parameter(Mandatory = $false)]
    [switch]$AllowRebootPassThru = $false,
    [Parameter(Mandatory = $false)]
    [switch]$TerminalServerMode = $false,
    [Parameter(Mandatory = $false)]
    [switch]$DisableLogging = $false
)

Try {
    ## Set the script execution policy for this process
    Try {
        Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop'
    } Catch {
    }

    ##*===============================================
    ##* VARIABLE DECLARATION
    ##*===============================================
    ## Variables: Application
    [String]$appVendor = ''
    [String]$appName = ''
    [String]$appVersion = ''
    [String]$appArch = ''
    [String]$appLang = 'EN'
    [String]$appRevision = '01'
    [String]$appScriptVersion = '1.0.0'
    [String]$appScriptDate = 'XX/XX/20XX'
    [String]$appScriptAuthor = '<author name>'
    ##*===============================================
    ## Variables: Install Titles (Only set here to override defaults set by the toolkit)
    [String]$installName = ''
    [String]$installTitle = ''

    ##* Do not modify section below
    #region DoNotModify

    ## Variables: Exit Code
    [Int32]$mainExitCode = 0

    ## Variables: Script
    [String]$deployAppScriptFriendlyName = 'Deploy Application'
    [Version]$deployAppScriptVersion = [Version]'3.10.1'
    [String]$deployAppScriptDate = '05/03/2024'
    [Hashtable]$deployAppScriptParameters = $PsBoundParameters

    ## Variables: Environment
    If (Test-Path -LiteralPath 'variable:HostInvocation') {
        $InvocationInfo = $HostInvocation
    } Else {
        $InvocationInfo = $MyInvocation
    }
    [String]$scriptDirectory = Split-Path -Path $InvocationInfo.MyCommand.Definition -Parent

    ## Dot source the required App Deploy Toolkit Functions
    Try {
        [String]$moduleAppDeployToolkitMain = "$scriptDirectory\AppDeployToolkit\AppDeployToolkitMain.ps1"
        If (-not (Test-Path -LiteralPath $moduleAppDeployToolkitMain -PathType 'Leaf')) {
            Throw "Module does not exist at the specified location [$moduleAppDeployToolkitMain]."
        }
        If ($DisableLogging) {
            . $moduleAppDeployToolkitMain -DisableLogging
        } Else {
            . $moduleAppDeployToolkitMain
        }
    } Catch {
        If ($mainExitCode -eq 0) {
            [Int32]$mainExitCode = 60008
        }
        Write-Error -Message "Module [$moduleAppDeployToolkitMain] failed to load: `n$($_.Exception.Message)`n `n$($_.InvocationInfo.PositionMessage)" -ErrorAction 'Continue'
        ## Exit the script, returning the exit code to SCCM
        If (Test-Path -LiteralPath 'variable:HostInvocation') {
            $script:ExitCode = $mainExitCode; Exit
        } Else {
            Exit $mainExitCode
        }
    }

    #endregion
    ##* Do not modify section above
    ##*===============================================
    ##* END VARIABLE DECLARATION
    ##*===============================================

    If ($deploymentType -ine 'Uninstall' -and $deploymentType -ine 'Repair') {
        ##*===============================================
        ##* PRE-INSTALLATION
        ##*===============================================
        [String]$installPhase = 'Pre-Installation'

        ## Show Welcome Message, close Internet Explorer if required, allow up to 3 deferrals, verify there is enough disk space to complete the install, and persist the prompt
        #Show-InstallationWelcome -CloseApps 'iexplore' -AllowDefer -DeferTimes 3 -CheckDiskSpace -PersistPrompt
        $targetOfficeExecutables = @(
            'lync',
            'winword',
            'excel',
            'msaccess', 
            'mstore',
            'infopath', 
            'setlang',
            'msouc',
            'ois',
            'onenote', 
            'outlook',
            'powerpnt', 
            'mspub',
            'groove', 
            'visio',
            'winproj', 
            'graph',
            'onedrive', 
            'teams',
            'ms-teams', 
            'olk' 
        )

        $foundOfficeExecutables = $false

        foreach ($exe in $targetOfficeExecutables) {
            $targetProcesses = @(Get-WmiObject -Query "Select * FROM Win32_Process WHERE Name='$exe'" -ErrorAction SilentlyContinue)
            if ($targetProcesses.Count -ne 0) {
                $foundOfficeExecutables = $true
                break
            }
        }

        if (-not $foundOfficeExecutables) {
            #No relevant Office processes found
    
        } else {
    
            #Office processes found
            Show-InstallationWelcome -CloseApps 'winword=Microsoft Office Word,excel=Microsoft Office Excel,POWERPNT=Microsoft Office PowerPoint,Outlook=Microsoft Office Outlook,ONENOTEM=Microsoft Office OneNote,MSACCESS=Microsoft Office Access,MSPUB=Microsoft Office Publisher,Visio=Microsoft Office Visio,WinProj=Microsoft Office Project' -BlockExecution -AllowDefer -DeferTimes 3 -CloseAppsCountdown 3600
            Write-Log -Message "Show-InstallationWelcome -CloseApps" -Source 'Office' -LogFileDirectory "C:\Temp\Win32app" -LogFileName "Office" -LogType 'CMTrace' 
        }
        ## Show Progress Message (with the default message)
        $officeFolderPath = Join-Path -ChildPath "office" -Path "C:\Temp"

        if (Test-Path $officeFolderPath) {
            Remove-Item -Path $officeFolderPath -Recurse -Force
        }

        # Create the Office folder path
        New-Item -ItemType Directory -Path $officeFolderPath -Force
        Write-Log -Message "Create folder C:\Temp\Office" -Source 'Office' -LogFileDirectory "C:\Temp\Win32app" -LogFileName "Office" -LogType 'CMTrace' 
        

        # Check if the Office installer file already exists
        $officeInstallerPath = Join-Path $officeFolderPath "setup.exe"
        if (-not (Test-Path $officeInstallerPath)) {
            # Download the Office 365 installer
            $officeInstallerDownloadURL = 'https://officecdn.microsoft.com/pr/wsus/setup.exe'
            try {
                Invoke-WebRequest -Uri $officeInstallerDownloadURL -OutFile $officeInstallerPath
                Write-Log -Message "Download Setup.exe" -Source 'Office' -LogFileDirectory "C:\Temp\Win32app" -LogFileName "Office" -LogType 'CMTrace' 
            } catch {
                Write-Host "Failed to download the Office installer: $_"
            }
        }

        # Download the Office 365 configuration file
        # PowerShell Script to Retrieve Azure AD Tenant Information from Registry

        # Function to retrieve a registry value
        function Get-RegistryValue {
            param (
                [string]$Path,
                [string]$Key
            )
    
            # Check if the registry path exists
            if (Test-Path $Path) {
                # Retrieve the value from the registry
                $value = Get-ItemProperty -Path $Path -Name $Key -ErrorAction SilentlyContinue | Select-Object -ExpandProperty $Key
        
                # Check if the value was found
                if ($value) {
                    return $value
                } else {
                    return "Value not found for $Key in $Path."
                }
            } else {
                return "Registry path does not exist: $Path."
            }
        }

        # Retrieve Azure AD Display Name
        try {
            $displayNamePath = "HKLM:\SYSTEM\ControlSet001\Control\CloudDomainJoin\TenantInfo\*"
            $displayNameKey = "DisplayName"
            $displayName = Get-RegistryValue -Path $displayNamePath -Key $displayNameKey
            Write-Output "Azure AD Display Name: $displayName"
        } catch {
            if ($_ -match "Registry path does not exist:") {
                $displayName = "Dummy Value"
                Write-Output "Azure AD Display Name: $displayName"
            } else {
                throw $_
            }
        }

        try {
            # Retrieve Azure AD Tenant ID
            $tenantIdPath = "HKLM:\SOFTWARE\Microsoft\Provisioning\Diagnostics\AutoPilot"
            $tenantIdKey = "TenantId"
            $tenantId = Get-RegistryValue -Path $tenantIdPath -Key $tenantIdKey
            Write-Output "Azure AD Tenant ID: $tenantId"
        } catch {
            if ($_ -match "Registry path does not exist:") {
                $tenantId = "Dummy Value"
                Write-Output "Azure AD Tenant ID: $tenantId"
            } else {
                throw $_
            }
        }

        $Microsoft365Appsforenterprise64bitonCurrentChannelUrl = "https://raw.githubusercontent.com/denniswesterman/M365-Apps/006880da9bf236522828e191b77570b85579eb07/Microsoft%20365%20Apps%20for%20enterprise%2064-bit%20on%20Current%20Channel.xml"
        $Microsoft365Appsforenterprise64bitonCurrentChannelPath = "Microsoft 365 Apps for enterprise 64-bit on Current Channel.xml"

        # Download config
        Invoke-WebRequest -Uri $Microsoft365Appsforenterprise64bitonCurrentChannelUrl -OutFile $officeFolderPath\$Microsoft365Appsforenterprise64bitonCurrentChannelPath -UseBasicParsing
        ((Get-Content -path $officeFolderPath\$Microsoft365Appsforenterprise64bitonCurrentChannelPath -Raw) -replace '<Company>', $displayName) | Set-Content -Path $officeFolderPath\$Microsoft365Appsforenterprise64bitonCurrentChannelPath
        ((Get-Content -path $officeFolderPath\$Microsoft365Appsforenterprise64bitonCurrentChannelPath -Raw) -replace '<tenantId>', $tenantId) | Set-Content -Path $officeFolderPath\$Microsoft365Appsforenterprise64bitonCurrentChannelPath
        $Download = Execute-Process -Path "$officeFolderPath\setup.exe" -Parameters "/download `"$officeFolderPath\Microsoft 365 Apps for enterprise 64-bit on Current Channel.xml`"" -Passthru
        Write-Log -Message "Download Office" -Source 'Office' -LogFileDirectory "C:\Temp\Win32app" -LogFileName "Office" -LogType 'CMTrace' 
        

        ## <Perform Pre-Installation tasks here>
        $processesExplorer = @(Get-CimInstance -ClassName 'Win32_Process' -Filter "Name like 'explorer.exe'" -ErrorAction 'Ignore')
        $esp = $false
        foreach ($processExplorer in $processesExplorer) {
            $user = (Invoke-CimMethod -InputObject $processExplorer -MethodName GetOwner).User
            if ($user -eq 'defaultuser0' -or $user -eq 'defaultuser1') { $esp = $true }
        }

        Write-Host $esp

        If ($esp -eq $True) {
            #In ESP - Do not Show Progress Banner
        } else {
            #Not in ESP - Show Progress Banner
            Show-InstallationProgress -StatusMessage "The latest version of Microsoft Apps 365 (64-bit) is being installed on your workstation. Thank you for your patience while we download and complete the setup. Once complete, this window will close."
            Write-Log -Message "Show-InstallationProgress -StatusMessage" -Source 'Office' -LogFileDirectory "C:\Temp\Win32app" -LogFileName "Office" -LogType 'CMTrace' 
        }
        

        ##*===============================================
        ##* INSTALLATION
        ##*===============================================
        [String]$installPhase = 'Installation'

        ## Handle Zero-Config MSI Installations
        If ($useDefaultMsi) {
            [Hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Install'; Path = $defaultMsiFile }; If ($defaultMstFile) {
                $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile)
            }
            Execute-MSI @ExecuteDefaultMSISplat; If ($defaultMspFiles) {
                $defaultMspFiles | ForEach-Object { Execute-MSI -Action 'Patch' -Path $_ }
            }
        }

        ## <Perform Installation tasks here>

        Set-location $officeFolderPath
        $Install = Execute-Process -Path "$officeFolderPath\setup.exe" -Parameters "/configure `"$officeFolderPath\Microsoft 365 Apps for enterprise 64-bit on Current Channel.xml`"" -Passthru
        Write-Log -Message "Exit" -Source 'Office' -LogFileDirectory "C:\Temp\Win32app" -LogFileName "Office" -LogType 'CMTrace' 
        # $ExitCode = $Install.exitcode
        # $ExitCode

        ##*===============================================
        ##* POST-INSTALLATION
        ##*===============================================
        [String]$installPhase = 'Post-Installation'

        ## <Perform Post-Installation tasks here>

        ## Display a message at the end of the install
        If (-not $useDefaultMsi) {
            Show-InstallationPrompt -Message 'You can customize text to appear at the end of an install or remove it completely for unattended installations.' -ButtonRightText 'OK' -Icon Information -NoWait
        }
    } ElseIf ($deploymentType -ieq 'Uninstall') {
        ##*===============================================
        ##* PRE-UNINSTALLATION
        ##*===============================================
        [String]$installPhase = 'Pre-Uninstallation'

        ## Show Welcome Message, close Internet Explorer with a 60 second countdown before automatically closing
        # Show-InstallationWelcome -CloseApps 'iexplore' -CloseAppsCountdown 60

        ## Show Progress Message (with the default message)
        Show-InstallationProgress

        ## <Perform Pre-Uninstallation tasks here>


        ##*===============================================
        ##* UNINSTALLATION
        ##*===============================================
        [String]$installPhase = 'Uninstallation'

        ## Handle Zero-Config MSI Uninstallations
        If ($useDefaultMsi) {
            [Hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Uninstall'; Path = $defaultMsiFile }; If ($defaultMstFile) {
                $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile)
            }
            Execute-MSI @ExecuteDefaultMSISplat
        }

        ## <Perform Uninstallation tasks here>


        ##*===============================================
        ##* POST-UNINSTALLATION
        ##*===============================================
        [String]$installPhase = 'Post-Uninstallation'

        ## <Perform Post-Uninstallation tasks here>


    } ElseIf ($deploymentType -ieq 'Repair') {
        ##*===============================================
        ##* PRE-REPAIR
        ##*===============================================
        [String]$installPhase = 'Pre-Repair'

        ## Show Welcome Message, close Internet Explorer with a 60 second countdown before automatically closing
        # Show-InstallationWelcome -CloseApps 'iexplore' -CloseAppsCountdown 60

        ## Show Progress Message (with the default message)
        Show-InstallationProgress

        ## <Perform Pre-Repair tasks here>

        ##*===============================================
        ##* REPAIR
        ##*===============================================
        [String]$installPhase = 'Repair'

        ## Handle Zero-Config MSI Repairs
        If ($useDefaultMsi) {
            [Hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Repair'; Path = $defaultMsiFile; }; If ($defaultMstFile) {
                $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile)
            }
            Execute-MSI @ExecuteDefaultMSISplat
        }
        ## <Perform Repair tasks here>

        ##*===============================================
        ##* POST-REPAIR
        ##*===============================================
        [String]$installPhase = 'Post-Repair'

        ## <Perform Post-Repair tasks here>


    }
    ##*===============================================
    ##* END SCRIPT BODY
    ##*===============================================

    ## Call the Exit-Script function to perform final cleanup operations
    Exit-Script -ExitCode $mainExitCode
} Catch {
    [Int32]$mainExitCode = 60001
    [String]$mainErrorMessage = "$(Resolve-Error)"
    Write-Log -Message $mainErrorMessage -Severity 3 -Source $deployAppScriptFriendlyName
    Show-DialogBox -Text $mainErrorMessage -Icon 'Stop'
    Exit-Script -ExitCode $mainExitCode
}
