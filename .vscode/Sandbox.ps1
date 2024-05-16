# Determine the path of the global script relative to this script
$globalScriptPath = Join-Path -Path (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent) -ChildPath 'Global.ps1'
if (-not (Test-Path -Path $globalScriptPath)) {
    Write-Error "The global script $globalScriptPath does not exist."
    exit
}
. $globalScriptPath

# Copy Resources
# Copy-Item -Path ".vscode\$LogonCommand" -Destination "$Win32App\" -Recurse -Force -Verbose -ErrorAction Ignore
robocopy "$ToolkitPath" "$Win32App" /E /R:3 /W:1 /NP

# Prepare Sandbox Configuration
$sandboxConfig = @"
<Configuration>
  <Networking>Enabled</Networking>
  <MappedFolders>
    <MappedFolder>
      <HostFolder>$Win32App</HostFolder>
      <SandboxFolder>$WDADesktop</SandboxFolder>
      <ReadOnly>true</ReadOnly>
    </MappedFolder>
  </MappedFolders>
  <LogonCommand>
    <Command>powershell -executionpolicy unrestricted -command `"Start-Process powershell -ArgumentList '-nologo -file $WDADesktop\$LogonCommand'`"</Command>
  </LogonCommand>
</Configuration>
"@

# Write the sandbox configuration to a file
$sandboxConfig | Out-File "$Win32App\$Application.wsb"

# Execute Sandbox
Start-Process explorer -ArgumentList "$Win32App\$Application.wsb" -Verbose -WindowStyle Maximized
