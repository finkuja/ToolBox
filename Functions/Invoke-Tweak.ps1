<#
.SYNOPSIS
    Applies or undoes various system tweaks.

.DESCRIPTION
    This function allows the user to apply or undo a set of predefined system tweaks. 
    It supports actions such as enabling detailed BSOD information, enabling God Mode, 
    and performing a clean boot, among others.

.PARAMETER Action
    Specifies the action to be taken. Valid values are "Apply" and "Undo".

.PARAMETER Tweak
    Specifies the tweak to be applied or undone. Valid values are:
    - CleanBoot
    - EnableDetailedBSODInformation
    - EnableGodMode
    - EnableClassicRightClickMenu
    - EnableEndTaskWithRightClick
    - ChangeIRPStackSize
    - ClipboardHistory
    - EnableVerboseLogonMessages
    - EnableVerboseStartupAndShutdownMessages

.PARAMETER window
    A System.Windows.Window object representing the parent window for message boxes.

.EXAMPLE
    Invoke-Tweak -Action Apply -Tweak EnableGodMode

    This command enables God Mode on the desktop.

.EXAMPLE
    Invoke-Tweak -Action Undo -Tweak CleanBoot

    This command undoes the clean boot by re-enabling previously disabled services.

.NOTES
    Author: Your Name
    Date: Today's Date
#>
function Invoke-Tweak {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet("Apply", "Undo")]
        [string]$Action,
        [string]$Tweak,
        [System.Windows.Window]$window
    )

    function Request-Reboot {
        # Prompt the user to reboot the system
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Changes have been applied. A system reboot is required for the changes to take effect. Do you want to reboot now?",
            "Reboot Required",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
    
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            Restart-Computer -Force
        }
        else {
            Write-Host "Reboot postponed by the user." -ForegroundColor Yellow
        }
    }

    Write-Host "$Action tweak...$Tweak" -ForegroundColor Cyan

    if ($Action -eq "Apply") {
        switch ($Tweak) {
            "CleanBoot" {
                # Add Apply logic for CleanBoot
                # Prompt the user
                $result = [System.Windows.Forms.MessageBox]::Show(
                    "This action will disable all non-Microsoft services. Do you want to continue?",
                    "Clean Boot",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
    
                if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                    # Perform a clean boot
                    Write-Host "Performing a clean boot..." -ForegroundColor Yellow
    
                    # Get all non-Microsoft services
                    $nonMicrosoftServices = Get-Service | Where-Object { $_.DisplayName -notmatch "^(Microsoft|Windows)" }
    
                    # Backup the services to a file
                    $backupFilePath = "$env:USERPROFILE\Documents\DisabledServicesBackup.txt"
                    $nonMicrosoftServices | Select-Object Name, DisplayName, Status | Export-Csv -Path $backupFilePath -NoTypeInformation
                    Write-Host "Backup of disabled services saved to $backupFilePath" -ForegroundColor Green
    
                    # Disable all non-Microsoft services
                    foreach ($service in $nonMicrosoftServices) {
                        try {
                            Set-Service -Name $service.Name -StartupType Disabled
                            Write-Host "Disabled service: $($service.DisplayName)" -ForegroundColor Green
                        }
                        catch {
                            Write-Host "Failed to disable service: $($service.DisplayName)" -ForegroundColor Red
                        }
                    }
    
                    # Disable all startup items using Task Scheduler
                    $startupTasks = Get-ScheduledTask | Where-Object { $_.TaskPath -notlike "\Microsoft\*" }
                    foreach ($task in $startupTasks) {
                        try {
                            Disable-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath
                            Write-Host "Disabled startup task: $($task.TaskName)" -ForegroundColor Green
                        }
                        catch {
                            Write-Host "Failed to disable startup task: $($task.TaskName)" -ForegroundColor Red
                        }
                    }
    
                    # Open System Configuration to verify changes
                    Start-Process "msconfig.exe" -ArgumentList "/4" -NoNewWindow -Wait
                }
                else {
                    Write-Host "Clean boot operation canceled by the user." -ForegroundColor Yellow
                }

            }
            "EnableDetailedBSODInformation" {
                # Add Apply logic for EnableDetailedBSODInformation
                # Enable detailed BSOD information
                Write-Host "Enabling detailed BSOD information..." -ForegroundColor Green
        
                try {
                    # Define the registry path
                    $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Control\CrashControl"
            
                    # Check if the registry path exists, and create it if it does not
                    if (-not (Test-Path $registryPath)) {
                        New-Item -Path $registryPath -Force
                        Write-Output "Created registry path: $registryPath"
                    }
            
                    # Set the registry key to enable detailed BSOD information
                    Set-ItemProperty -Path $registryPath -Name "DisplayParameters" -Value 1 -Force
                    Write-Host "Detailed BSOD information has been enabled." -ForegroundColor Green
                }
                catch {
                    Write-Host "Failed to enable detailed BSOD information: $_" -ForegroundColor Red
                }
            }
            "EnableGodMode" {
                # Add Apply logic for EnableGodMode
                $desktopPath = [System.Environment]::GetFolderPath('Desktop')
                $godModePath = "$desktopPath\GodMode.{ED7BA470-8E54-465E-825C-99712043E01C}"
                if (-not (Test-Path $godModePath)) {
                    try {
                        New-Item -Path $godModePath -ItemType Directory -Force | Out-Null
                        Write-Host "God Mode has been enabled on the desktop." -ForegroundColor Green
                    }
                    catch {
                        Write-Host "Failed to enable God Mode. Error: $_" -ForegroundColor Red
                    }
                }
                else {
                    Write-Host "God Mode is already enabled on the desktop." -ForegroundColor Yellow
                }
            }
            "EnableClassicRightClickMenu" {
                # Add Apply logic for EnableClassicRightClickMenu
                # Enable Classic Right Click Menu
                New-Item -Path "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}" -Name "InprocServer32" -Force -Value ""
                Write-Host "Classic Right Click Menu has been enabled." -ForegroundColor Green

                # Restart explorer.exe
                Write-Host "Restarting explorer.exe ..." -ForegroundColor Green
                $process = Get-Process -Name "explorer"
                Stop-Process -InputObject $process
            }
            "EnableEndTaskWithRightClick" {
                # Add Apply logic for EnableEndTaskWithRightClick
                try {
                    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarDeveloperSettings"
                    $regName = "TaskbarEndTask"
                    $regValue = 1

                    Write-Host "Applying EnableEndTaskWithRightClick..."

                    # Ensure the registry path exists
                    if (-not (Test-Path $regPath)) {
                        New-Item -Path $regPath -Force | Out-Null
                        Write-Host "Created registry path: $regPath"
                    }

                    # Set the registry value, creating it if it doesn't exist
                    New-ItemProperty -Path $regPath -Name $regName -Value $regValue -Force | Out-Null
                    Write-Host "Set registry value: $regName to $regValue at $regPath" -ForegroundColor Green
                    Write-Host "End Task with Rigth Click has been enabled." -ForegroundColor Green
                }
                catch {
                    Write-Host "Failed to apply EnableEndTaskWithRightClick: $_" -ForegroundColor Red
                }
            }
            "ChangeIRPStackSize" {
                # Add Apply logic for ChangeIRPStackSize
                try {
                    # Define the registry path
                    $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters"
                                
                    # Check if the registry path exists, and create it if it does not
                    if (-not (Test-Path $registryPath)) {
                        New-Item -Path $registryPath -Force
                        Write-Output "Created registry path: $registryPath"
                    }
                                
                    # Set the IRPStackSize to 32
                    Set-ItemProperty -Path $registryPath -Name "IRPStackSize" -Value 32 -Type DWord -Force
                    Write-Host "IRPStackSize has been set to 32." -ForegroundColor Green
                    Request-Reboot
                }
                catch {
                    Write-Host "Failed to set IRPStackSize: $_" -ForegroundColor Red
                }
            }
            "ClipboardHistory" {
                # Add Apply logic for ClipboardHistory
                # Add Apply logic for ClipboardHistory
                try {
                    # Define the registry path
                    $registryPath = "HKCU:\Software\Microsoft\Clipboard"
    
                    # Check if the registry path exists, and create it if it does not
                    if (-not (Test-Path $registryPath)) {
                        New-Item -Path $registryPath -Force
                        Write-Output "Created registry path: $registryPath"
                    }
    
                    # Set the EnableClipboardHistory to 1
                    Set-ItemProperty -Path $registryPath -Name "EnableClipboardHistory" -Value 1 -Type DWord -Force
                    Write-Host "Clipboard History has been enabled." -ForegroundColor Green
                }
                catch {
                    Write-Host "Failed to enable Clipboard History: $_" -ForegroundColor Red
                }
            }
            "EnableVerboseLogonMessages" {
                # Enable verbose logon messages
                Write-Output "Enabling verbose logon messages..."
                # Add your code here
                try {
                    # Define the registry path
                    $registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
            
                    # Check if the registry path exists, and create it if it does not
                    if (-not (Test-Path $registryPath)) {
                        New-Item -Path $registryPath -Force
                        Write-Output "Created registry path: $registryPath"
                    }
            
                    # Set the registry key to enable verbose logon messages
                    Set-ItemProperty -Path $registryPath -Name "VerboseStatus" -Value 1 -Force
                    Write-Host "Verbose logon messages have been enabled." -ForegroundColor Green
                }
                catch {
                    Write-Host "Failed to enable verbose logon messages: $_" -ForegroundColor Red
                }
            }
            "EnableVerboseStartupAndShutdownMessages" {
                # Add Apply logic for EnableVerboseStartupAndShutdownMessages
                try {
                    # Define the registry path
                    $registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
                                
                    # Check if the registry path exists, and create it if it does not
                    if (-not (Test-Path $registryPath)) {
                        New-Item -Path $registryPath -Force
                        Write-Output "Created registry path: $registryPath"
                    }
                                
                    # Set the VerboseStatus to 1
                    Set-ItemProperty -Path $registryPath -Name "VerboseStatus" -Value 1 -Type DWord -Force
                    Write-Host "Verbose startup and shutdown messages have been enabled." -ForegroundColor Green
                }
                catch {
                    Write-Host "Failed to enable verbose startup and shutdown messages: $_" -ForegroundColor Red
                }
            }
            default {
                Write-Error "Unknown checkbox tag: $($control.Tag)"
            }
        }
    }
    elseif ($Action -eq "Undo") {
        switch ($Tweak) {
            "CleanBoot" {
                # Add Undo logic for CleanBoot
                # Undo clean boot
                Write-Host "Undoing clean boot..." -ForegroundColor Yellow

                # Path to the backup file
                $backupFilePath = "$env:USERPROFILE\Documents\DisabledServicesBackup.txt"

                # Check if the backup file exists
                if (Test-Path $backupFilePath) {
                    try {
                        # Read the backup file
                        $disabledServices = Import-Csv -Path $backupFilePath

                        # Re-enable the services
                        foreach ($service in $disabledServices) {
                            try {
                                Set-Service -Name $service.Name -StartupType Automatic
                                Write-Host "Re-enabled service: $($service.DisplayName)" -ForegroundColor Green
                            }
                            catch {
                                Write-Host "Failed to re-enable service: $($service.DisplayName)" -ForegroundColor Red
                            }
                        }

                        # Remove the backup file after re-enabling services
                        Remove-Item -Path $backupFilePath -Force
                        Write-Host "Clean boot undo completed." -ForegroundColor Green
                    }
                    catch {
                        Write-Host "Failed to undo clean boot. Error: $_" -ForegroundColor Red
                    }
                }
                else {
                    Write-Host "No backup file found for clean boot undo." -ForegroundColor Red
                }
            }
            "EnableDetailedBSODInformation" {
                # Add Undo logic for EnableDetailedBSODInformation
                # Disable detailed BSOD information
                Write-Host "Disabling detailed BSOD information..." -ForegroundColor Green
    
                try {
                    # Define the registry path
                    $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Control\CrashControl"
        
                    # Check if the registry path exists
                    if (Test-Path $registryPath) {
                        # Set the registry key to disable detailed BSOD information
                        Set-ItemProperty -Path $registryPath -Name "DisplayParameters" -Value 0 -Force
                        Write-Host "Detailed BSOD information has been disabled." -ForegroundColor Green
                    }
                    else {
                        Write-Host "Registry path does not exist: $registryPath" -ForegroundColor Red
                    }
                }
                catch {
                    Write-Host "Failed to disable detailed BSOD information: $_" -ForegroundColor Red
                }
            }
            "EnableGodMode" {
                # Add Undo logic for EnableGodMode
                # Disable God Mode
                $desktopPath = [System.Environment]::GetFolderPath('Desktop')
                $godModePath = "$desktopPath\GodMode.{ED7BA470-8E54-465E-825C-99712043E01C}"
                if (Test-Path $godModePath) {
                    try {
                        Remove-Item -Path $godModePath -Recurse -Force
                        Write-Host "God Mode has been disabled." -ForegroundColor Green
                    }
                    catch {
                        Write-Host "Failed to disable God Mode. Error: $_" -ForegroundColor Red
                    }
                }
                else {
                    Write-Host "God Mode folder does not exist." -ForegroundColor Yellow
                }
            }
            "EnableClassicRightClickMenu" {
                # Add Undo logic for EnableClassicRightClickMenu
                # Disable Classic Right Click Menu
                Remove-Item -Path "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}" -Recurse -Confirm:$false -Force
                Write-Host "Classic Right Click Menu has been disabled." -ForegroundColor Green

                # Restart explorer.exe
                Write-Host "Restarting explorer.exe ..." -ForegroundColor Green
                $process = Get-Process -Name "explorer"
                Stop-Process -InputObject $process
            }
            "EnableEndTaskWithRightClick" {
                # Add Undo logic for EnableEndTaskWithRightClick
                # Remove registry key to disable right click end task
                $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarDeveloperSettings"
                $regName = "TaskbarEndTask"
                $regValue = 0
        
                #Ensure the registry path exists
                if (-not (Test-Path $regPath)) {
                    New-Item -Path $regPath -Force | Out-Null
                }
                #Remove the registry value, creating it if it doesn't exist
                New-ItemProperty -Path $regPath -Name $regName -Value $regValue -Force | Out-Null
                Write-Host "End Task with Rigth Click has been disabled." -ForegroundColor Green
            }
            "ChangeIRPStackSize" {
                # Add Undo logic for ChangeIRPStackSize
                # Add Undo logic for ChangeIRPStackSize
                try {
                    # Define the registry path
                    $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters"
        
                    # Check if the registry path exists
                    if (Test-Path $registryPath) {
                        # Remove the IRPStackSize entry
                        Remove-ItemProperty -Path $registryPath -Name "IRPStackSize" -Force
                        Write-Host "IRPStackSize has been removed." -ForegroundColor Green
                        # Prompt the user to reboot the system
                        Request-Reboot
                    }
                    else {
                        Write-Host "Registry path does not exist: $registryPath" -ForegroundColor Yellow
                    }
                }
                catch {
                    Write-Host "Failed to remove IRPStackSize: $_" -ForegroundColor Red
                }
                      
            }
            "ClipboardHistory" {
                # Add Undo logic for ClipboardHistory
                try {
                    # Define the registry path
                    $registryPath = "HKCU:\Software\Microsoft\Clipboard"
        
                    # Check if the registry path exists
                    if (Test-Path $registryPath) {
                        # Set the EnableClipboardHistory back to 0
                        Set-ItemProperty -Path $registryPath -Name "EnableClipboardHistory" -Value 0 -Type DWord -Force
                        Write-Host "Clipboard History has been disabled." -ForegroundColor Green
                    }
                    else {
                        Write-Host "Registry path does not exist: $registryPath" -ForegroundColor Yellow
                    }
                }
                catch {
                    Write-Host "Failed to disable Clipboard History: $_" -ForegroundColor Red
                }
            }
            "EnableVerboseLogonMessages" {
                # Add Undo logic for EnableVerboseLogonMessages
                # Disable verbose logon messages
                Write-Output "Disabling verbose logon messages..."
                # Add your code here
                try {
                    # Define the registry path
                    $registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
            
                    # Check if the registry path exists
                    if (Test-Path $registryPath) {
                        # Remove the registry key to disable verbose logon messages
                        Remove-ItemProperty -Path $registryPath -Name "VerboseStatus" -Force
                        Write-Host "Verbose logon messages have been disabled." -ForegroundColor Green
                    }
                    else {
                        Write-Host "Registry path does not exist: $registryPath" -ForegroundColor Red
                    }
                }
                catch {
                    Write-Host "Failed to disable verbose logon messages: $_" -ForegroundColor Red
                }
            }
            "EnableVerboseStartupAndShutdownMessages" {
                # Add Undo logic for EnableVerboseStartupAndShutdownMessages
                try {
                    # Define the registry path
                    $registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
                                
                    # Check if the registry path exists
                    if (Test-Path $registryPath) {
                        # Set the VerboseStatus back to 0
                        Set-ItemProperty -Path $registryPath -Name "VerboseStatus" -Value 0 -Type DWord -Force
                        Write-Host "Verbose startup and shutdown messages have been disabled." -ForegroundColor Green
                    }
                    else {
                        Write-Host "Registry path does not exist: $registryPath" -ForegroundColor Yellow
                    }
                }
                catch {
                    Write-Host "Failed to disable verbose startup and shutdown messages: $_" -ForegroundColor Red
                }
            }
            default {
                Write-Error "Unknown checkbox tag: $($control.Tag)"
            }
        }
    }
    else {
        Write-Error "Invalid action specified."
    }
            
}