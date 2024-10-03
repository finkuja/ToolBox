# Ensure you run this script with administrator privileges


Write-Output " ______  _____  _____    _______          _   ____ 
|  ____|/ ____|/ ____|  |__   __|        | | |  _ \             
| |__  | (___ | (___       | | ___   ___ | | | |_) | ___  __  __
|  __|  \___ \ \___ \      | |/ _ \ / _ \| | |  _ < / _ \ \ \/ /
| |____ ____) |____) |     | | (_) | (_) | | | |_) | (_) | |  |
|______|_____/|_____/      |_|\___/ \___/|_| |____/ \___/ /_/\_\
 
 === Version Alpha 0.1 ===

 === Created by: Carlos Alvarez Magariños ===
 "
# Import necessary .NET assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Check if the script is running with administrator privileges
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    Write-Host "The ESSToolBox needs to run as Administrator, trying to elevate the permissions..." -ForegroundColor Yellow
    
    # Get the current script content
    $scriptContent = (irm aka.ms/esstoolbox)
    
    # Define a temporary file path
    $tempFilePath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "ESSToolBox.ps1")
    
    # Save the script content to the temporary file
    [System.IO.File]::WriteAllText($tempFilePath, $scriptContent)
    
    # Create a new process to run the script with administrator privileges
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = "powershell.exe"
    $startInfo.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$tempFilePath`""
    $startInfo.Verb = "runas"
    
    try {
        # Start the new process
        $process = [System.Diagnostics.Process]::Start($startInfo)
        $process.WaitForExit()
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to run the script with administrator privileges. Please run this script with administrator privileges manually.")
    }
    
    # Exit the current script
    exit
} else {
    Write-Host "Running with administrator privileges." -ForegroundColor Green
}

#Check if winget is installed
Write-Host "Checking if Windows Package Manager (winget) is installed..."
$winget = Get-Command winget -ErrorAction SilentlyContinue
if ($null -eq $winget) {
    [System.Windows.Forms.MessageBox]::Show("Windows Package Manager (winget) is not installed. Please install it from https://github.com/microsoft/winget-cli/releases")
    Start-Process "https://github.com/microsoft/winget-cli/releases"
    exit
}

# Get the winget source list and find the 'msstore' source
$wingetSource = & winget source list | Where-Object { $_.Name -eq 'msstore' }

# Check if the 'msstore' source exists and if the source agreement is not accepted
if ($wingetSource) {
    if (-not $wingetSource.Accepted) {
        # Update the 'msstore' source and accept the source agreements
        Start-Process "winget" -ArgumentList "source update --name msstore" -NoNewWindow -Wait
        Write-Host "MS Store Source Agreement has been accepted." -ForegroundColor Green
    } else {
        Write-Host "MS Store Source Agreement is already accepted." -ForegroundColor Green
    }
} else {
    Write-Host "MS Store source not found." -ForegroundColor Yellow
}

# Create the form of the main GUI window
$form = New-Object System.Windows.Forms.Form
$form.Text = "ESS Tool Box a.01"
$form.Size = New-Object System.Drawing.Size(500, 400)
$form.StartPosition = "CenterScreen"

# Create a TabControl
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Size = New-Object System.Drawing.Size(480, 350)
$tabControl.Location = New-Object System.Drawing.Point(10, 10)
$form.Controls.Add($tabControl)

# Create the Install tab
$tabInstall = New-Object System.Windows.Forms.TabPage
$tabInstall.Text = "Install / Update"
$tabControl.Controls.Add($tabInstall)

# Define column positions
$column1X = 20
$column2X = 200

# Create checkboxes for packages in the Install tab

$checkboxAdobe = New-Object System.Windows.Forms.CheckBox
$checkboxAdobe.Text = "Adobe Reader"
$checkboxAdobe.Name = "Adobe"
$checkboxAdobe.AutoSize = $true
$checkboxAdobe.Location = New-Object System.Drawing.Point($column1X, 20)
$tabInstall.Controls.Add($checkboxAdobe)

$checkboxAdobeCloud = New-Object System.Windows.Forms.CheckBox
$checkboxAdobeCloud.Text = "Adobe Cloud"
$checkboxAdobeCloud.Name = "Adobe.Cloud"
$checkboxAdobeCloud.AutoSize = $true
$checkboxAdobeCloud.Location = New-Object System.Drawing.Point($column1X, 50)
$tabInstall.Controls.Add($checkboxAdobeCloud)

$checkboxOffice = New-Object System.Windows.Forms.CheckBox
$checkboxOffice.Text = "Microsoft Office 365"
$checkboxOffice.Name = "Microsoft.Office"
$checkboxOffice.AutoSize = $true
$checkboxOffice.Location = New-Object System.Drawing.Point($column1X, 80)
$tabInstall.Controls.Add($checkboxOffice)

$checkboxOneNote = New-Object System.Windows.Forms.CheckBox
$checkboxOneNote.Text = "Microsoft OneNote"
$checkboxOneNote.Name = "Microsoft.OneNote"
$checkboxOneNote.AutoSize = $true
$checkboxOneNote.Location = New-Object System.Drawing.Point($column1X, 110)
$tabInstall.Controls.Add($checkboxOneNote)

$checkboxTeams = New-Object System.Windows.Forms.CheckBox
$checkboxTeams.Text = "Microsoft Teams"
$checkboxTeams.Name = "Microsoft.Teams"
$checkboxTeams.AutoSize = $true
$checkboxTeams.Location = New-Object System.Drawing.Point($column1X, 140)
$tabInstall.Controls.Add($checkboxTeams)

$checkboxNetFrameworks = New-Object System.Windows.Forms.CheckBox
$checkboxNetFrameworks.Text = ".NET Frameworks"
$checkboxNetFrameworks.Name = "NetFrameworks"
$checkboxNetFrameworks.AutoSize = $true
$checkboxNetFrameworks.Location = New-Object System.Drawing.Point($column1X, 170)
$tabInstall.Controls.Add($checkboxNetFrameworks)

$checkboxPowerAutomate = New-Object System.Windows.Forms.CheckBox
$checkboxPowerAutomate.Text = "Power Automate"
$checkboxPowerAutomate.Name = "PowerAutomate"
$checkboxPowerAutomate.AutoSize = $true
$checkboxPowerAutomate.Location = New-Object System.Drawing.Point($column1X, 200)
$tabInstall.Controls.Add($checkboxPowerAutomate)

$checkboxPowerToys = New-Object System.Windows.Forms.CheckBox
$checkboxPowerToys.Text = "PowerToys"
$checkboxPowerToys.Name = "Microsoft.PowerToys"
$checkboxPowerToys.AutoSize = $true
$checkboxPowerToys.Location = New-Object System.Drawing.Point($column2X, 20)
$tabInstall.Controls.Add($checkboxPowerToys)

$checkboxQuickAssist = New-Object System.Windows.Forms.CheckBox
$checkboxQuickAssist.Text = "Quick Assist"
$checkboxQuickAssist.Name = "QuickAssist"
$checkboxQuickAssist.AutoSize = $true
$checkboxQuickAssist.Location = New-Object System.Drawing.Point($column2X, 50)
$tabInstall.Controls.Add($checkboxQuickAssist)

$checkboxSurfaceDiagnosticToolkit = New-Object System.Windows.Forms.CheckBox
$checkboxSurfaceDiagnosticToolkit.Text = "Surface Diagnostic Toolkit"
$checkboxSurfaceDiagnosticToolkit.Name = "SurfaceDiagnosticToolkit"
$checkboxSurfaceDiagnosticToolkit.AutoSize = $true
$checkboxSurfaceDiagnosticToolkit.Location = New-Object System.Drawing.Point($column2X, 80)
$tabInstall.Controls.Add($checkboxSurfaceDiagnosticToolkit)

$checkboxVisio = New-Object System.Windows.Forms.CheckBox
$checkboxVisio.Text = "Visio"
$checkboxVisio.Name = "Visio"
$checkboxVisio.AutoSize = $true
$checkboxVisio.Location = New-Object System.Drawing.Point($column2X, 110)
$tabInstall.Controls.Add($checkboxVisio)

# Create an Install button in the Install tab
$buttonInstall = New-Object System.Windows.Forms.Button
$buttonInstall.Text = "Install"
$buttonInstall.AutoSize = $true
$buttonInstall.Location = New-Object System.Drawing.Point(20, 290)
$tabInstall.Controls.Add($buttonInstall)

# Define the action for the Install button
$buttonInstall.Add_Click({
    if ($checkboxOffice.Checked) {
        Start-Process "winget" -ArgumentList "install --id Microsoft.Office -e" -NoNewWindow -Wait
    }
    if ($checkboxPowerToys.Checked) {
        Start-Process "winget" -ArgumentList "install --id Microsoft.PowerToys -e" -NoNewWindow -Wait
    }
    if ($checkboxTeams.Checked) {
        Start-Process "winget" -ArgumentList "install --id Microsoft.Teams -e" -NoNewWindow -Wait
    }
    if ($checkboxOneNote.Checked) {
        Start-Process "winget" -ArgumentList "install --id Microsoft.OneNote -e" -NoNewWindow -Wait
    }
    if ($checkboxOffice.Checked) {
        Start-Process "winget" -ArgumentList "install --id Microsoft.Office -e" -NoNewWindow -Wait
    }
    if ($checkboxPowerToys.Checked) {
        Start-Process "winget" -ArgumentList "install --id Microsoft.PowerToys -e" -NoNewWindow -Wait
    }
    if ($checkboxTeams.Checked) {
        Start-Process "winget" -ArgumentList "install --id Microsoft.Teams -e" -NoNewWindow -Wait
    }
    if ($checkboxOneNote.Checked) {
        Start-Process "winget" -ArgumentList "install --id Microsoft.OneNote -e" -NoNewWindow -Wait
    }
    if ($checkboxAdobe.Checked) {
        Start-Process "winget" -ArgumentList "install --id Adobe -e" -NoNewWindow -Wait
    }
    if ($checkboxAdobeCloud.Checked) {
        Start-Process "winget" -ArgumentList "install --id Adobe.Cloud -e" -NoNewWindow -Wait
    }
    if ($checkboxPowerAutomate.Checked) {
        Start-Process "winget" -ArgumentList "install --id PowerAutomate -e" -NoNewWindow -Wait
    }
    if ($checkboxVisio.Checked) {
        Start-Process "winget" -ArgumentList "install --id Visio -e" -NoNewWindow -Wait
    }
    if ($checkboxNetFrameworks.Checked) {
        Start-Process "winget" -ArgumentList "install --id NetFrameworks -e" -NoNewWindow -Wait
    }
    if ($checkboxQuickAssist.Checked) {
        Start-Process "winget" -ArgumentList "install --id QuickAssist -e" -NoNewWindow -Wait
    }
    if ($checkboxSurfaceDiagnosticToolkit.Checked) {
        Start-Process "winget" -ArgumentList "install --id SurfaceDiagnosticToolkit -e" -NoNewWindow -Wait
    }
    [System.Windows.Forms.MessageBox]::Show("Selected packages have been installed.")
})

# Create an Uninstall button in the Install tab
$buttonUninstall = New-Object System.Windows.Forms.Button
$buttonUninstall.Text = "Uninstall"
$buttonUninstall.AutoSize = $true
$buttonUninstall.Location = New-Object System.Drawing.Point(120, 290)
$tabInstall.Controls.Add($buttonUninstall)

# Define the action for the Uninstall button
$buttonUninstall.Add_Click({
    if ($checkboxOffice.Checked) {
        Start-Process "winget" -ArgumentList "uninstall --id Microsoft.Office -e" -NoNewWindow -Wait
    }
    if ($checkboxPowerToys.Checked) {
        Start-Process "winget" -ArgumentList "uninstall --id Microsoft.PowerToys -e" -NoNewWindow -Wait
    }
    if ($checkboxTeams.Checked) {
        Start-Process "winget" -ArgumentList "uninstall --id Microsoft.Teams -e" -NoNewWindow -Wait
    }
    if ($checkboxOneNote.Checked) {
        Start-Process "winget" -ArgumentList "uninstall --id Microsoft.OneNote -e" -NoNewWindow -Wait
    }
    if ($checkboxAdobe.Checked) {
        Start-Process "winget" -ArgumentList "uninstall --id Adobe -e" -NoNewWindow -Wait
    }
    if ($checkboxAdobeCloud.Checked) {
        Start-Process "winget" -ArgumentList "uninstall --id Adobe.Cloud -e" -NoNewWindow -Wait
    }
    if ($checkboxPowerAutomate.Checked) {
        Start-Process "winget" -ArgumentList "uninstall --id PowerAutomate -e" -NoNewWindow -Wait
    }
    if ($checkboxVisio.Checked) {
        Start-Process "winget" -ArgumentList "uninstall --id Visio -e" -NoNewWindow -Wait
    }
    if ($checkboxNetFrameworks.Checked) {
        Start-Process "winget" -ArgumentList "uninstall --id NetFrameworks -e" -NoNewWindow -Wait
    }
    if ($checkboxQuickAssist.Checked) {
        Start-Process "winget" -ArgumentList "uninstall --id QuickAssist -e" -NoNewWindow -Wait
    }
    if ($checkboxSurfaceDiagnosticToolkit.Checked) {
        Start-Process "winget" -ArgumentList "uninstall --id SurfaceDiagnosticToolkit -e" -NoNewWindow -Wait
    }
    [System.Windows.Forms.MessageBox]::Show("Selected packages have been uninstalled.")
})

# Create a Get Installed Packages button in the Install tab
$buttonGetPackages = New-Object System.Windows.Forms.Button
$buttonGetPackages.Text = "Get Installed"
$buttonGetPackages.AutoSize = $true
$buttonGetPackages.Location = New-Object System.Drawing.Point(220, 290)
$tabInstall.Controls.Add($buttonGetPackages)

# Define the action for the Get Packages button
$buttonGetPackages.Add_Click({
    # Run the winget list command and capture the output directly
    $output = & winget list
    
    # Split the output into lines and skip the first two lines (headers)
    $outputLines = $output -split '\r?\n'
    
    # Iterate through each control in the install tab
    foreach ($control in $tabInstall.Controls) {
        if ($control -is [System.Windows.Forms.CheckBox]) {
            $checkboxName = $control.Name
            # Check if the checkbox name is present in the output lines
            if ($outputLines -match $checkboxName) {
                $control.Checked = $true
            } else {
                $control.Checked = $false
            }
        }
    }
})

# Create the Tweak tab
$tabTweak = New-Object System.Windows.Forms.TabPage
$tabTweak.Text = "Tweak"
$tabControl.Controls.Add($tabTweak)

# Create controls for the Tweak tab
$checkboxRightClickEndTask = New-Object System.Windows.Forms.CheckBox
$checkboxRightClickEndTask.Text = "Enable end task with right click"
$checkboxRightClickEndTask.Name = "EnableRightClickEndTask"
$checkboxRightClickEndTask.AutoSize = $true
$checkboxRightClickEndTask.Location = New-Object System.Drawing.Point(20, 20)
$tabTweak.Controls.Add($checkboxRightClickEndTask)

$checkboxRunDiskCleanup = New-Object System.Windows.Forms.CheckBox
$checkboxRunDiskCleanup.Text = "Run disk cleanup"
$checkboxRunDiskCleanup.Name = "RunDiskCleanup"
$checkboxRunDiskCleanup.AutoSize = $true
$checkboxRunDiskCleanup.Location = New-Object System.Drawing.Point(20, 50)
$tabTweak.Controls.Add($checkboxRunDiskCleanup)

$checkboxDetailedBSOD = New-Object System.Windows.Forms.CheckBox
$checkboxDetailedBSOD.Text = "Enable detailed BSOD information"
$checkboxDetailedBSOD.Name = "EnableDetailedBSOD"
$checkboxDetailedBSOD.AutoSize = $true
$checkboxDetailedBSOD.Location = New-Object System.Drawing.Point(20, 80)
$tabTweak.Controls.Add($checkboxDetailedBSOD)

$checkboxVerboseLogon = New-Object System.Windows.Forms.CheckBox
$checkboxVerboseLogon.Text = "Enable verbose logon messages"
$checkboxVerboseLogon.Name = "EnableVerboseLogon"
$checkboxVerboseLogon.AutoSize = $true
$checkboxVerboseLogon.Location = New-Object System.Drawing.Point(20, 110)
$tabTweak.Controls.Add($checkboxVerboseLogon)

$checkboxDeleteTempFiles = New-Object System.Windows.Forms.CheckBox
$checkboxDeleteTempFiles.Text = "Delete temporary files"
$checkboxDeleteTempFiles.Name = "DeleteTempFiles"
$checkboxDeleteTempFiles.AutoSize = $true
$checkboxDeleteTempFiles.Location = New-Object System.Drawing.Point(20, 140)
$tabTweak.Controls.Add($checkboxDeleteTempFiles)

$buttonApply = New-Object System.Windows.Forms.Button
$buttonApply.Text = "Apply"
$buttonApply.Location = New-Object System.Drawing.Point(20, 190)
$tabTweak.Controls.Add($buttonApply)

$buttonUndo = New-Object System.Windows.Forms.Button
$buttonUndo.Text = "Undo"
$buttonUndo.Location = New-Object System.Drawing.Point(120, 190)
$tabTweak.Controls.Add($buttonUndo)

# Define the action for the Apply button
$buttonApply.Add_Click({
    if ($checkboxRightClickEndTask.Checked) {
        # Add registry key to enable right click end task
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarDeveloperSettings"
        $regName = "TaskbarEndTask"
        $regValue = 1
        #Ensure the registry path exists
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        #Set the registry value, creating it if it doesn't exist
        New-ItemProperty -Path $regPath -Name $regName -Value $regValue -Force | Out-Null
    }
    if ($checkboxRunDiskCleanup.Checked) {
        # Run disk cleanup
        Start-Process "cleanmgr.exe" -ArgumentList "/sagerun:1" -NoNewWindow -Wait
        
        # Check the Windows version
        $osVersion = [System.Environment]::OSVersion.Version
        $majorVersion = $osVersion.Major
        #$minorVersion = $osVersion.Minor
        $buildNumber = $osVersion.Build
    
        # Run DISM command with appropriate arguments based on Windows version
        if ($majorVersion -ge 10 -and $buildNumber -ge 14393) {
            # Windows 10 version 1607 (Anniversary Update) or later
            Start-Process "dism.exe" -ArgumentList "/Online /Cleanup-Image /ResetBase" -NoNewWindow -Wait
        } else {
            # Older versions of Windows
            Start-Process "dism.exe" -ArgumentList "/Online /Cleanup-Image /StartComponentCleanup" -NoNewWindow -Wait
        }
    }
    if ($checkboxDetailedBSOD.Checked) {
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
        } catch {
            Write-Host "Failed to enable detailed BSOD information: $_" -ForegroundColor Red
        }
    }
    if ($checkboxVerboseLogon.Checked) {
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
        } catch {
            Write-Host "Failed to enable verbose logon messages: $_" -ForegroundColor Red
        }
    }
    if ($checkboxDeleteTempFiles.Checked) {
        # Delete temporary files
        Write-Host "Deleting temporary files..." -ForegroundColor Green

        # Check if the C:\Windows\Temp path exists
        if (Test-Path "C:\Windows\Temp") {
            Get-ChildItem -Path "C:\Windows\Temp" *.* -Recurse | Remove-Item -Force -Recurse
        } else {
            Write-Host "Path C:\Windows\Temp does not exist." -ForegroundColor Red
        }

        # Check if the $env:TEMP path exists
        if (Test-Path $env:TEMP) {
            Get-ChildItem -Path $env:TEMP *.* -Recurse | Remove-Item -Force -Recurse
        } else {
            Write-Host "Path $env:TEMP does not exist." -ForegroundColor Red
        }

    }
    [System.Windows.Forms.MessageBox]::Show("Selected tweaks have been applied.")
})

# Define the action for the Undo button
$buttonUndo.Add_Click({
    if ($checkboxRightClickEndTask.Checked) {
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
    }
    if ($checkboxRunDiskCleanup.Checked) {
        # Undo disk cleanup
        Write-Output "Nothing to do here..."
    }
    if ($checkboxDetailedBSOD.Checked) {
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
            } else {
                Write-Host "Registry path does not exist: $registryPath" -ForegroundColor Red
            }
        } catch {
        Write-Host "Failed to disable detailed BSOD information: $_" -ForegroundColor Red
        }
    }
    if ($checkboxVerboseLogon.Checked) {
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
            } else {
                Write-Host "Registry path does not exist: $registryPath" -ForegroundColor Red
            }
        } catch {
            Write-Host "Failed to disable verbose logon messages: $_" -ForegroundColor Red
        }
    }
    if ($checkboxDeleteTempFiles.Checked) {
        # Undo delete temporary files
        Write-Output "Nothing to do here..."
        # Add your code here
    }
    [System.Windows.Forms.MessageBox]::Show("Selected tweaks have been undone.")
})

# Create the Fix tab
$tabFix = New-Object System.Windows.Forms.TabPage
$tabFix.Text = "Fix"
$tabControl.Controls.Add($tabFix)

# Create controls for the Fix tab

# Apps section
$sectionApps = New-Object System.Windows.Forms.GroupBox
$sectionApps.Text = "Apps"
$sectionApps.Size = New-Object System.Drawing.Size(480, 100)
$sectionApps.Location = New-Object System.Drawing.Point($column1X, 20)
$tabFix.Controls.Add($sectionApps)

# Create a button to remove Adobe Cloud
$buttonRemoveAdobeCloud = New-Object System.Windows.Forms.Button
$buttonRemoveAdobeCloud.Text = "Remove Adobe Cloud"
$buttonRemoveAdobeCloud.AutoSize = $true
$buttonRemoveAdobeCloud.Location = New-Object System.Drawing.Point($column1X, 30)
$buttonRemoveAdobeCloud.Add_Click({
    # Remove Adobe Cloud
    Write-Output "Removing Adobe Cloud..."
    # Add your code here
})
$sectionApps.Controls.Add($buttonRemoveAdobeCloud)

# Create a button to remove Adobe Reader
$buttonRemoveAdobeReader = New-Object System.Windows.Forms.Button
$buttonRemoveAdobeReader.Text = "Remove Adobe Reader"
$buttonRemoveAdobeReader.AutoSize = $true
$buttonRemoveAdobeReader.Location = New-Object System.Drawing.Point($column1X, 60)
$buttonRemoveAdobeReader.Add_Click({
    # Remove Adobe Reader
    Write-Output "Removing Adobe Reader..."
    # Add your code here
})
$sectionApps.Controls.Add($buttonRemoveAdobeReader)

# System section
$sectionSystem = New-Object System.Windows.Forms.GroupBox
$sectionSystem.Text = "System"
$sectionSystem.Size = New-Object System.Drawing.Size(480, 100)
$sectionSystem.Location = New-Object System.Drawing.Point(20, 140)
$tabFix.Controls.Add($sectionSystem)

# Create a button to reset Windows Update
$buttonResetWinUpdate = New-Object System.Windows.Forms.Button
$buttonResetWinUpdate.Text = "Reset Windows Update"
$buttonResetWinUpdate.AutoSize = $true
$buttonResetWinUpdate.Location = New-Object System.Drawing.Point($column1X, 30)
$buttonResetWinUpdate.Add_Click({
    # Reset Windows Update
    Write-Output "Resetting Windows Update..."
    # Add your code here
})
$sectionSystem.Controls.Add($buttonResetWinUpdate)

# Create a button to reset network
$buttonResetNetwork = New-Object System.Windows.Forms.Button
$buttonResetNetwork.Text = "Reset Network"
$buttonResetNetwork.AutoSize = $true
$buttonResetNetwork.Location = New-Object System.Drawing.Point($column1X, 60)
$buttonResetNetwork.Add_Click({
    # Reset network
    Write-Output "Resetting network..."
    # Add your code here
})
$sectionSystem.Controls.Add($buttonResetNetwork)

# Create a button to run system repair
$buttonSysRepair = New-Object System.Windows.Forms.Button
$buttonSysRepair.Text = "System Repair"
$buttonSysRepair.AutoSize = $true
$buttonSysRepair.Location = New-Object System.Drawing.Point($column2X, 30)
$buttonSysRepair.Add_Click({
    # Run system repair
    Write-Output "Running system repair..."
    # Add your code here
})
$sectionSystem.Controls.Add($buttonSysRepair)

# Office Apps section
$sectionOfficeApps = New-Object System.Windows.Forms.GroupBox
$sectionOfficeApps.Text = "Office Apps"
$sectionOfficeApps.Size = New-Object System.Drawing.Size(480, 100)
$sectionOfficeApps.Location = New-Object System.Drawing.Point($column1X, 260)
$tabFix.Controls.Add($sectionOfficeApps)

# Create a button to rebuild .OST file
$buttonRebuildOST = New-Object System.Windows.Forms.Button
$buttonRebuildOST.Text = "Rebuild .OST file"
$buttonRebuildOST.AutoSize = $true
$buttonRebuildOST.Location = New-Object System.Drawing.Point($column1X, 30)
$buttonRebuildOST.Add_Click({
    # Rebuild .OST file
    Write-Output "Rebuilding .OST file..."
    # Add your code here
})
$sectionOfficeApps.Controls.Add($buttonRebuildOST)

# ...

# Run the form
$form.ShowDialog()
