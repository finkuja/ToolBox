function Invoke-OfficeRepair {

    # Inform the user that the repair process is starting
    $result = [System.Windows.Forms.MessageBox]::Show(
        "This action will start the Office repair process. Please save any important work before proceeding. Do you want to continue?", 
        "Confirmation", 
        [System.Windows.Forms.MessageBoxButtons]::YesNo, 
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host "Repairing Office..." -ForegroundColor Green

        # Check for running Office applications
        $officeApps = @("WINWORD", "EXCEL", "POWERPNT", "OUTLOOK", "ONENOTE", "MSACCESS", "MSPUB", "VISIO", "LYNC")
        $runningApps = @()
        foreach ($app in $officeApps) {
            $process = Get-Process -Name $app -ErrorAction SilentlyContinue
            if ($process) {
                $runningApps += $app
            }
        }

        if ($runningApps.Count -gt 0) {
            $appsList = $runningApps -join ", "
            $closeResult = [System.Windows.Forms.MessageBox]::Show(
                "The following Office applications are currently running: $appsList. Do you want to close them and proceed with the repair?", 
                "Office Applications Running", 
                [System.Windows.Forms.MessageBoxButtons]::OKCancel, 
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            if ($closeResult -eq [System.Windows.Forms.DialogResult]::OK) {
                foreach ($app in $runningApps) {
                    try {
                        Stop-Process -Name $app -Force
                        Write-Host "$app has been closed." -ForegroundColor Green
                    }
                    catch {
                        Write-Host "Failed to close $app. Office repair process has been canceled." -ForegroundColor Red
                        return
                    }
                }
            }
            else {
                Write-Host "Office repair process was canceled by the user." -ForegroundColor Yellow
                return
            }
        }

        # Create a custom form to prompt the user to choose between App Reset and Repair
        $choiceForm = New-Object System.Windows.Forms.Form
        $choiceForm.Text = "Choose Action"
        $choiceForm.Size = New-Object System.Drawing.Size(300, 150)
        $choiceForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

        $choiceLabel = New-Object System.Windows.Forms.Label
        $choiceLabel.Text = "Do you want to perform an Office?"
        $choiceLabel.AutoSize = $true
        $choiceLabel.Location = New-Object System.Drawing.Point(10, 10)
        $choiceForm.Controls.Add($choiceLabel)

        $repairButton = New-Object System.Windows.Forms.Button
        $repairButton.Text = "Repair"
        $repairButton.Location = New-Object System.Drawing.Point(10, 50)
        $repairButton.Add_Click({
                $choiceForm.Tag = "Repair"
                $choiceForm.Close()
            })
        $choiceForm.Controls.Add($repairButton)

        $resetButton = New-Object System.Windows.Forms.Button
        $resetButton.Text = "Reset"
        $resetButton.Location = New-Object System.Drawing.Point(100, 50)
        $resetButton.Add_Click({
                $choiceForm.Tag = "Reset"
                $choiceForm.Close()
            })
        $choiceForm.Controls.Add($resetButton)

        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.Location = New-Object System.Drawing.Point(190, 50)
        $cancelButton.Add_Click({
                $choiceForm.Tag = "Cancel"
                $choiceForm.Close()
            })
        $choiceForm.Controls.Add($cancelButton)

        $choiceForm.ShowDialog()

        $choiceResult = $choiceForm.Tag
        if ($choiceResult -eq "Repair") {
            # Perform Repair
            try {
                # Start the Office repair process using Windows 11 reset and repair app feature
                $progressForm = New-Object System.Windows.Forms.Form
                $progressForm.Text = "Repairing Office"
                $progressForm.Size = New-Object System.Drawing.Size(400, 100)
                $progressForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
                $progressLabel = New-Object System.Windows.Forms.Label
                $progressLabel.Text = "Repairing Office, please wait..."
                $progressLabel.AutoSize = $true
                $progressLabel.Location = New-Object System.Drawing.Point(10, 10)
                $progressForm.Controls.Add($progressLabel)
                $progressForm.Show()

                # Use the Windows 11 reset and repair app feature
                Start-Process -FilePath "ms-settings:appsfeatures-app" -ArgumentList "Microsoft Office" -Wait

                $progressForm.Close()
                Write-Host "Office repair process has been started." -ForegroundColor Green
            }
            catch {
                Write-Host "Failed to start the Office repair process." -ForegroundColor Red
            }
        }
        elseif ($choiceResult -eq "Reset") {
            # Perform App Reset
            try {
                # Start the Office reset process using Windows 11 reset and repair app feature
                $progressForm = New-Object System.Windows.Forms.Form
                $progressForm.Text = "Resetting Office"
                $progressForm.Size = New-Object System.Drawing.Size(400, 100)
                $progressForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
                $progressLabel = New-Object System.Windows.Forms.Label
                $progressLabel.Text = "Resetting Office, please wait..."
                $progressLabel.AutoSize = $true
                $progressLabel.Location = New-Object System.Drawing.Point(10, 10)
                $progressForm.Controls.Add($progressLabel)
                $progressForm.Show()

                # Use the Windows 11 reset and repair app feature
                Start-Process -FilePath "ms-settings:appsfeatures-app" -ArgumentList "Microsoft Office" -Wait

                $progressForm.Close()
                Write-Host "Office reset process has been started." -ForegroundColor Green
            }
            catch {
                Write-Host "Failed to start the Office reset process." -ForegroundColor Red
            }
        }
        else {
            Write-Host "Office repair/reset process was canceled by the user." -ForegroundColor Yellow
            return
        }

        [System.Windows.Forms.MessageBox]::Show(
            "The Office repair/reset process is complete.", 
            "Information", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
}