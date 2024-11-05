# Function to uninstall Edge by changing the region to Ireland and uninstalling Edge, then changing it back From Chris Titus Tech  winutils.ps1 script
<#
    .SYNOPSIS
    This will uninstall Edge by changing the region to Ireland and uninstalling Edge, then changing it back.

    .DESCRIPTION
    The Uninstall-EdgeBrowser function stops any running instances of Microsoft Edge and Widgets, changes the system region to Ireland, and then uninstalls Microsoft Edge. 
    After the uninstallation, it restores the original region settings.

    .NOTES
    Author: Chris Titus Tech github.com/ChrisTitusTech/winutil/blob/main/docs/dev/features/Fixes/Uninstall-EdgeBrowser.md
    File Name: Uninstall-EdgeBrowser.ps1
    The function uses the Get-Process, Stop-Process, Get-ItemProperty, Remove-ItemProperty, and Start-Process cmdlets to uninstall Microsoft Edge.
    #>
function Uninstall-EdgeBrowser {

    $msedgeProcess = Get-Process -Name "msedge" -ErrorAction SilentlyContinue
    $widgetsProcess = Get-Process -Name "widgets" -ErrorAction SilentlyContinue
    
    # Checking if Microsoft Edge is running
    if ($msedgeProcess) {
        Stop-Process -Name "msedge" -Force
    }
    else {
        Write-Output "msedge process is not running."
    }
    
    # Checking if Widgets is running
    if ($widgetsProcess) {
        Stop-Process -Name "widgets" -Force
    }
    else {
        Write-Output "widgets process is not running."
    }
    
    function Uninstall-Process {
        <#
            .SYNOPSIS
            Uninstalls a process by modifying registry settings and executing the uninstall command.
    
            .PARAMETER Key
            The registry key associated with the process to be uninstalled.
    
            .DESCRIPTION
            This function temporarily changes the system region to Ireland, modifies necessary registry settings, and executes the uninstall command for the specified process. After uninstallation, it restores the original region settings and registry permissions.
    
            .PARAMETER Key
            The registry key associated with the process to be uninstalled.
    
            .NOTES
            Author: Chris Titus Tech
            #>
        param(
            [Parameter(Mandatory = $true)]
            [string]$Key
        )
        $originalNation = [microsoft.win32.registry]::GetValue('HKEY_USERS\.DEFAULT\Control Panel\International\Geo', 'Nation', [Microsoft.Win32.RegistryValueKind]::String)
        # Set Nation to 84 (Ireland) temporarily
        [microsoft.win32.registry]::SetValue('HKEY_USERS\.DEFAULT\Control Panel\International\Geo', 'Nation', 68, [Microsoft.Win32.RegistryValueKind]::String) | Out-Null
        # credits to he3als for the Acl commands
        $fileName = "IntegratedServicesRegionPolicySet.json"
        $pathISRPS = [Environment]::SystemDirectory + "\" + $fileName
        $aclISRPS = Get-Acl -Path $pathISRPS
        $aclISRPSBackup = [System.Security.AccessControl.FileSecurity]::new()
        $aclISRPSBackup.SetSecurityDescriptorSddlForm($acl.Sddl)
        if (Test-Path -Path $pathISRPS) {
            try {
                $admin = [System.Security.Principal.NTAccount]$(New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')).Translate([System.Security.Principal.NTAccount]).Value
                $aclISRPS.SetOwner($admin)
                $rule = New-Object System.Security.AccessControl.FileSystemAccessRule($admin, 'FullControl', 'Allow')
                $aclISRPS.AddAccessRule($rule)
                Set-Acl -Path $pathISRPS -AclObject $aclISRPS
                Rename-Item -Path $pathISRPS -NewName ($fileName + '.bak') -Force
            }
            catch {
                Write-Error "Failed to set owner for $pathISRPS"
            }
        }
        $baseKey = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate'
        $registryPath = $baseKey + '\ClientState\' + $Key
        if (!(Test-Path -Path $registryPath)) {
            Write-Host "Registry key not found: $registryPath"
            return
        }
        Remove-ItemProperty -Path $registryPath -Name "experiment_control_labels" -ErrorAction SilentlyContinue | Out-Null
        $uninstallString = (Get-ItemProperty -Path $registryPath).UninstallString
        $uninstallArguments = (Get-ItemProperty -Path $registryPath).UninstallArguments
        if ([string]::IsNullOrEmpty($uninstallString) -or [string]::IsNullOrEmpty($uninstallArguments)) {
            Write-Host "Cannot find uninstall methods for $Mode"
            return
        }
        $uninstallArguments += " --force-uninstall --delete-profile"
        if (!(Test-Path -Path $uninstallString)) {
            Write-Host "setup.exe not found at: $uninstallString"
            return
        }
        Start-Process -FilePath $uninstallString -ArgumentList $uninstallArguments -Wait -NoNewWindow -Verbose
        # Restore Acl
        if (Test-Path -Path ($pathISRPS + '.bak')) {
            Rename-Item -Path ($pathISRPS + '.bak') -NewName $fileName -Force
            Set-Acl -Path $pathISRPS -AclObject $aclISRPSBackup
        }
        # Restore Nation
        [microsoft.win32.registry]::SetValue('HKEY_USERS\.DEFAULT\Control Panel\International\Geo', 'Nation', $originalNation, [Microsoft.Win32.RegistryValueKind]::String) | Out-Null
        if ((Get-ItemProperty -Path $baseKey).IsEdgeStableUninstalled -eq 1) {
            Write-Host "Edge Stable has been successfully uninstalled"
        }
    }
}