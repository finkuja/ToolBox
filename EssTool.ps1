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

# Create an Install button
$buttonInstall = New-Object System.Windows.Forms.Button
$buttonInstall.Text = "Install"
$buttonInstall.Location = New-Object System.Drawing.Point(20, 100)
$form.Controls.Add($buttonInstall)

# Define the action for the Install button
$buttonInstall.Add_Click({
    if ($checkboxOffice.Checked) {
        Start-Process "winget" -ArgumentList "install --id Microsoft.Office -e" -NoNewWindow -Wait
    }
    if ($checkboxPowerToys.Checked) {
        Start-Process "winget" -ArgumentList "install --id Microsoft.PowerToys -e" -NoNewWindow -Wait
    }
    [System.Windows.Forms.MessageBox]::Show("Selected packages have been installed.")
})

# Run the form
$form.ShowDialog()
