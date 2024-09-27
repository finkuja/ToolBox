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
$checkboxOffice.Location = New-Object System.Drawing.Point(20, 20)
$tabInstall.Controls.Add($checkboxOffice)

$checkboxPowerToys = New-Object System.Windows.Forms.CheckBox
$checkboxPowerToys.Text = "PowerToys"
$checkboxPowerToys.Location = New-Object System.Drawing.Point(20, 50)
$tabInstall.Controls.Add($checkboxPowerToys)

$checkboxTeams = New-Object System.Windows.Forms.CheckBox
$checkboxTeams.Text = "Microsoft Teams"
$checkboxTeams.Location = New-Object System.Drawing.Point(20, 80)
$tabInstall.Controls.Add($checkboxTeams)

$checkboxOneNote = New-Object System.Windows.Forms.CheckBox
$checkboxOneNote.Text = "Microsoft OneNote"
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
# ...

# Run the form
$form.ShowDialog()
