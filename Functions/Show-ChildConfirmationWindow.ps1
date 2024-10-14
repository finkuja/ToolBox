function Show-ChildConfirmationWindow {
    param (
        [string]$xamlPath,
        [string]$message,
        [string]$title = "Confirmation",
        [string[]]$comboBoxItems = @()
    )

    if (Test-Path $xamlPath) {
        try {
            $childXaml = Get-Content -Path $xamlPath -Raw
            $childReader = (New-Object System.Xml.XmlTextReader (New-Object System.IO.StringReader $childXaml))
            $childWindow = [Windows.Markup.XamlReader]::Load($childReader)

            # Set the window title
            $childWindow.Title = $title

            # Find the message TextBlock and set its text
            $messageTextBlock = $childWindow.FindName("MessageTextBlock")
            if ($null -ne $messageTextBlock) {
                $messageTextBlock.Text = $message
            }

            # Find the ComboBox and load items if provided
            $comboBox = $childWindow.FindName("OptionsComboBox")
            if ($null -ne $comboBox) {
                if ($comboBoxItems.Count -gt 0) {
                    $comboBox.Visibility = [System.Windows.Visibility]::Visible
                    foreach ($item in $comboBoxItems) {
                        $comboBox.Items.Add($item)
                    }
                }
                else {
                    $comboBox.Visibility = [System.Windows.Visibility]::Collapsed
                }
            }

            # Find the Yes and No buttons and add click event handlers
            $yesButton = $childWindow.FindName("YesButton")
            $noButton = $childWindow.FindName("NoButton")

            if ($null -ne $yesButton) {
                $yesButton.Add_Click({
                        $childWindow.DialogResult = $true
                        $childWindow.Close()
                    })
            }

            if ($null -ne $noButton) {
                $noButton.Add_Click({
                        $childWindow.DialogResult = $false
                        $childWindow.Close()
                    })
            }

            # Show the window and return the result
            $result = $childWindow.ShowDialog()
            return $result, $comboBox.SelectedItem
        }
        catch {
            Write-Host "Failed to load child window XAML: $_" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Error: File not found - $xamlPath" -ForegroundColor Red
    }
}