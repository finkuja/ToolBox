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

# Create checkboxes for packages
$checkboxOffice = New-Object System.Windows.Forms.CheckBox
$checkboxOffice.Text = "Microsoft Office 365"
$checkboxOffice.Location = New-Object System.Drawing.Point(20, 20)
$form.Controls.Add($checkboxOffice)

$checkboxPowerToys = New-Object System.Windows.Forms.CheckBox
$checkboxPowerToys.Text = "PowerToys"
$checkboxPowerToys.Location = New-Object System.Drawing.Point(20, 50)
$form.Controls.Add($checkboxPowerToys)

$checkboxTeams = New-Object System.Windows.Forms.CheckBox
$checkboxTeams.Text = "Microsoft Teams"
$checkboxTeams.Location = New-Object System.Drawing.Point(20, 80)
$form.Controls.Add($checkboxTeams)

$checkboxOneNote = New-Object System.Windows.Forms.CheckBox
$checkboxOneNote.Text = "Microsoft OneNote"
$checkboxOneNote.Location = New-Object System.Drawing.Point(20, 110)
$form.Controls.Add($checkboxOneNote)

# Create an Install button
$buttonInstall = New-Object System.Windows.Forms.Button
$buttonInstall.Text = "Install"
$buttonInstall.Location = New-Object System.Drawing.Point(20, 160)
$form.Controls.Add($buttonInstall)

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

# Run the form
$form.ShowDialog()
