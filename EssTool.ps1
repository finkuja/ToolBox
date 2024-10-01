# Ensure you run this script with administrator privileges
# Import necessary .NET assemblies

Write-Output "  ______  _____  _____    _______          _ 
 |  ____|/ ____|/ ____|  |__   __|        | |
 | |__  | (___ | (___       | | ___   ___ | |
 |  __|  \___ \ \___ \      | |/ _ \ / _ \| |
 | |____ ____) |____) |     | | (_) | (_) | |
 |______|_____/|_____/      |_|\___/ \___/|_|"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "ESS Tool Box a.01"
$form.Size = New-Object System.Drawing.Size(400, 300)
$form.StartPosition = "CenterScreen"

# Create a TabControl
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Size = New-Object System.Drawing.Size(380, 250)
$tabControl.Location = New-Object System.Drawing.Point(10, 10)
$form.Controls.Add($tabControl)

# Create the Install tab
$tabInstall = New-Object System.Windows.Forms.TabPage
$tabInstall.Text = "Install / Update"
$tabControl.Controls.Add($tabInstall)

# Create checkboxes for packages in the Install tab
$checkboxOffice = New-Object System.Windows.Forms.CheckBox
$checkboxOffice.Text = "Microsoft Office 365"
$checkboxOffice.Name = "Microsoft.Office"
$checkboxOffice.Location = New-Object System.Drawing.Point(20, 20)
$tabInstall.Controls.Add($checkboxOffice)

$checkboxPowerToys = New-Object System.Windows.Forms.CheckBox
$checkboxPowerToys.Text = "PowerToys"
$checkboxPowerToys.Name = "Microsoft.PowerToys"
$checkboxPowerToys.Location = New-Object System.Drawing.Point(20, 50)
$tabInstall.Controls.Add($checkboxPowerToys)

$checkboxTeams = New-Object System.Windows.Forms.CheckBox
$checkboxTeams.Text = "Microsoft Teams"
$checkboxTeams.Name = "Microsoft.Teams"
$checkboxTeams.Location = New-Object System.Drawing.Point(20, 80)
$tabInstall.Controls.Add($checkboxTeams)

$checkboxOneNote = New-Object System.Windows.Forms.CheckBox
$checkboxOneNote.Text = "Microsoft OneNote"
$checkboxOneNote.Name = "Microsoft.OneNote"
$checkboxOneNote.Location = New-Object System.Drawing.Point(20, 110)
$tabInstall.Controls.Add($checkboxOneNote)

# Create an Install button in the Install tab
$buttonInstall = New-Object System.Windows.Forms.Button
$buttonInstall.Text = "Install"
$buttonInstall.Location = New-Object System.Drawing.Point(20, 160)
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
    [System.Windows.Forms.MessageBox]::Show("Selected packages have been installed.")
})

# Create a Get Packages button in the Install tab
$buttonGetPackages = New-Object System.Windows.Forms.Button
$buttonGetPackages.Text = "Get Packages"
$buttonGetPackages.Location = New-Object System.Drawing.Point(120, 160)
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

# ...

# Create the Fix tab
$tabFix = New-Object System.Windows.Forms.TabPage
$tabFix.Text = "Fix"
$tabControl.Controls.Add($tabFix)

# Create controls for the Fix tab
$sectionTeams = New-Object System.Windows.Forms.GroupBox
$sectionTeams.Text = "Teams"
$sectionTeams.Size = New-Object System.Drawing.Size(350, 150)
$sectionTeams.Location = New-Object System.Drawing.Point(20, 20)
$tabFix.Controls.Add($sectionTeams)

$buttonFixOutlookAddin = New-Object System.Windows.Forms.Button
$buttonFixOutlookAddin.Text = "Fix Outlook Addin"
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
})# ..$sectionTeams.Controls.Add($buttonFixOutlookAddin)

# Run the form
$form.ShowDialog()
