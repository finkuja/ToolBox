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

#Check if winget is installed
$winget = Get-Command winget -ErrorAction SilentlyContinue
if ($winget -eq $null) {
    [System.Windows.Forms.MessageBox]::Show("Windows Package Manager (winget) is not installed. Please install it from https://github.com/microsoft/winget-cli/releases")
    Start-Process "https://github.com/microsoft/winget-cli/releases"
    exit
}

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
$checkboxAdobe.Text = "Adobe"
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
$buttonInstall.Location = New-Object System.Drawing.Point(20, 230)
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

# Create a Get Installed Packages button in the Install tab
$buttonGetPackages = New-Object System.Windows.Forms.Button
$buttonGetPackages.Text = "Get Installed"
$buttonGetPackages.AutoSize = $true
$buttonGetPackages.Location = New-Object System.Drawing.Point(120, 230)
$tabInstall.Controls.Add($buttonGetPackages)

# Define the action for the Get Packages button
$buttonGetPackages.Add_Click({
    $output = Start-Process "winget" -ArgumentList "list" -NoNewWindow -Wait -PassThru
    $installedPackages = $output.StandardOutput -split '\r?\n' | Select-Object -Skip 2 | ForEach-Object { $_ -split '\s{2,}' } | Where-Object { $_[0] -ne '' }
    
    foreach ($package in $installedPackages) {
        $packageId = $package[1]
        Write-Line $package[1]
        foreach ($control in $tabInstall.Controls) {
            if ($control.Name -eq $packageId) {
                $control.Checked = $true
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
$checkboxRightClickEndTask.Text = "Enable right click end task"
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

$buttonApply = New-Object System.Windows.Forms.Button
$buttonApply.Text = "Apply"
$buttonApply.Location = New-Object System.Drawing.Point(20, 100)
$tabTweak.Controls.Add($buttonApply)

$buttonUndo = New-Object System.Windows.Forms.Button
$buttonUndo.Text = "Undo"
$buttonUndo.Location = New-Object System.Drawing.Point(120, 100)
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
        Start-Process "dism.exe" -ArgumentList "/Online /Cleanup-Image /ResetBase" -NoNewWindow -Wait
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
    [System.Windows.Forms.MessageBox]::Show("Selected tweaks have been undone.")
})
# ...

# Create the Fix tab
$tabFix = New-Object System.Windows.Forms.TabPage
$tabFix.Text = "Fix"
$tabControl.Controls.Add($tabFix)

# Create controls for the Fix tab

# Teams section
$sectionTeams = New-Object System.Windows.Forms.GroupBox
$sectionTeams.Text = "Teams"
$sectionTeams.Size = New-Object System.Drawing.Size(350, 150)
$sectionTeams.Location = New-Object System.Drawing.Point(20, 20)
$tabFix.Controls.Add($sectionTeams)

# Create a button to fix the Teams Outlook Add-in
$buttonFixOutlookAddin = New-Object System.Windows.Forms.Button
$buttonFixOutlookAddin.Text = "Fix Outlook Addin"
$buttonFixOutlookAddin.AutoSize = $true
$buttonFixOutlookAddin.Location = New-Object System.Drawing.Point(20, 30)
$buttonFixOutlookAddin.Add_Click({
    Stop-Process -Name "Teams" -Force
    Stop-Process -Name "Outlook" -Force
    $currentUser = $env:USERNAME
    $squirrelTempPath = "C:\Users\$currentUser\AppData\Local\SquirrelTemp"
    $teamsPath = "C:\Users\$currentUser\AppData\Local\Microsoft\Teams"
    if (Test-Path $squirrelTempPath) {
        Remove-Item -Path $squirrelTempPath -Recurse -Force
    }

    if (Test-Path $teamsPath) {
        Remove-Item -Path $teamsPath -Recurse -Force
    }

    Start-Process "winget" -ArgumentList "uninstall --id Microsoft.Teams -e" -NoNewWindow -Wait
    Start-Process "winget" -ArgumentList "install --id Microsoft.Teams -e" -NoNewWindow -Wait

    [System.Windows.Forms.MessageBox]::Show("Teams Outlook Add-in has been fixed.")
})
$sectionTeams.Controls.Add($buttonFixOutlookAddin)

# Run the form
$form.ShowDialog()
