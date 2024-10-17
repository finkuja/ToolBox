<#
.SYNOPSIS
Rebuilds the Outlook OST file.

.DESCRIPTION
The Invoke-RebuildOST function rebuilds the Outlook Offline Storage Table (OST) file. 
It provides feedback to the user by displaying a message before starting the rebuild process.

.PARAMETER None
This function does not take any parameters.

.NOTES
File Name      : Invoke-RebuildOST.ps1
The function uses the Start-Process cmdlet to launch Outlook with the resetfolders and cleanreminders switches.
The function is part of the ToolBox project and is stored in the GitHub repository https://github.com/finkuja/ToolBox

.EXAMPLE
Invoke-RebuildOST
Prompts the user and runs the Outlook OST rebuild process.

#>
function Invoke-RebuildOST {
    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to proceed with rebuilding the OST files? This will delete or rename selected files.",
        "Confirmation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirmation -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host "Operation cancelled by the user."
        return
    }

    # Check if Outlook and Teams are running
    $outlookRunning = Get-Process -Name Outlook -ErrorAction SilentlyContinue
    $teamsRunning = @()
    $teamsRunning += Get-Process -Name Teams -ErrorAction SilentlyContinue
    $teamsRunning += Get-Process -Name ms-teams -ErrorAction SilentlyContinue
    $teamsRunning = $teamsRunning | Where-Object { $_ -ne $null }

    # Terminate Outlook and Teams
    if ($outlookRunning) {
        Stop-Process -Name Outlook -Force
    }
    foreach ($teamProcess in $teamsRunning) {
        Stop-Process -Id $teamProcess.Id -Force
    }

    # Path to the Outlook .ost files
    $ostPath = "$env:LOCALAPPDATA\Microsoft\Outlook"

    # Check if the path exists
    if (Test-Path $ostPath) {
        # Function to load files into the CheckedListBox
        function Get-Files {
            param (
                [System.Windows.Forms.CheckedListBox]$checkedListBox,
                [bool]$showBackups
            )
            $checkedListBox.Items.Clear()
            $filter = if ($showBackups) { "*.bak" } else { "*.ost" }
            $files = Get-ChildItem -Path $ostPath -Filter $filter -Recurse -ErrorAction SilentlyContinue
            foreach ($file in $files) {
                $checkedListBox.Items.Add($file.FullName)
            }
        }

        # Create a form to display the list of .ost files
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Select .OST Files"
        $form.Size = New-Object System.Drawing.Size(800, 440)  # Set a wider form size
        $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $form.MaximizeBox = $false  # Disable the maximize button
        $form.MinimizeBox = $false  # Optionally, disable the minimize button

        $label = New-Object System.Windows.Forms.Label
        $label.Text = "Select the .ost files to delete or rename:"
        $label.AutoSize = $true
        $label.Location = New-Object System.Drawing.Point(10, 10)
        $form.Controls.Add($label)

        $checkedListBox = New-Object System.Windows.Forms.CheckedListBox
        $checkedListBox.Size = New-Object System.Drawing.Size(760, 280)  # Adjust the size to fit within the form
        $checkedListBox.Location = New-Object System.Drawing.Point(10, 40)
        $form.Controls.Add($checkedListBox)

        $seeBackupsCheckbox = New-Object System.Windows.Forms.CheckBox
        $seeBackupsCheckbox.Text = "Show .bak files"
        $seeBackupsCheckbox.AutoSize = $true
        $seeBackupsCheckbox.Location = New-Object System.Drawing.Point(280, 330)
        $seeBackupsCheckbox.Add_CheckedChanged({
                Get-Files -checkedListBox $checkedListBox -showBackups $seeBackupsCheckbox.Checked
            })
        $form.Controls.Add($seeBackupsCheckbox)

        $deleteButton = New-Object System.Windows.Forms.Button
        $deleteButton.Text = "Delete"
        $deleteButton.AutoSize = $true
        $deleteButton.Location = New-Object System.Drawing.Point(10, 330)
        $deleteButton.Add_Click({
                $deletedFiles = $checkedListBox.CheckedItems | ForEach-Object {
                    Remove-Item -Path $_ -Force
                    $_
                }
                if ($seeBackupsCheckbox.Checked) {
                    Write-Host "Selected .bak files have been deleted:" -ForegroundColor Green
                }
                else {
                    Write-Host "Selected .ost files have been deleted:" -ForegroundColor Green
                }
                $deletedFiles | ForEach-Object { Write-Host $_ -ForegroundColor Green }
                $form.Close()
            })
        $form.Controls.Add($deleteButton)

        $renameButton = New-Object System.Windows.Forms.Button
        $renameButton.Text = "Rename"
        $renameButton.AutoSize = $true
        $renameButton.Location = New-Object System.Drawing.Point(100, 330)
        $renameButton.Add_Click({
                $renamedFiles = $checkedListBox.CheckedItems | ForEach-Object {
                    $newName = "$($_.FullName).bak"
                    Rename-Item -Path $_ -NewName $newName
                    $newName
                }
                if ($seeBackupsCheckbox.Checked) {
                    Write-Host "Selected .bak files have been renamed:" -ForegroundColor Green
                }
                else {
                    Write-Host "Selected .ost files have been renamed:" -ForegroundColor Green
                }
                $renamedFiles | ForEach-Object { Write-Host $_ -ForegroundColor Green }
                $form.Close()
            })
        $form.Controls.Add($renameButton)

        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.AutoSize = $true
        $cancelButton.Location = New-Object System.Drawing.Point(190, 330)
        $cancelButton.Add_Click({
                $form.Close()
            })
        $form.Controls.Add($cancelButton)

        # Load initial .ost files
        Get-Files -checkedListBox $checkedListBox -showBackups $false

        $form.ShowDialog()
    }
    else {
        Write-Host "The path to .ost files does not exist." -ForegroundColor Red
    }

    # Reopen Outlook and Teams if they were running before
    if ($outlookRunning) {
        Start-Process "Outlook"
    }
    foreach ($teamProcess in $teamsRunning) {
        Start-Process $teamProcess.Name
    }

    [System.Windows.Forms.MessageBox]::Show("The process is complete. Outlook and Teams have been restarted if they were running before.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}