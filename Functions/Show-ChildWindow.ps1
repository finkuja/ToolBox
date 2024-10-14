function Show-ChildWindow {
    param (
        [string]$xamlPath
    )

    if (Test-Path $xamlPath) {
        try {
            $childXaml = Get-Content -Path $xamlPath -Raw
            $childReader = (New-Object System.Xml.XmlTextReader (New-Object System.IO.StringReader $childXaml))
            $childWindow = [Windows.Markup.XamlReader]::Load($childReader)
            
            # Find the close button and add a click event handler
            $closeButton = $childWindow.FindName("CloseButton")
            if ($null -ne $closeButton) {
                $closeButton.Add_Click({
                        $childWindow.Close()
                    })
            }

            # Find all buttons and assign click event handlers based on their tags
            $buttons = $childWindow.FindName("ChildWindowControlPanel").Children | Where-Object { $_ -is [System.Windows.Controls.Button] }
            foreach ($button in $buttons) {
                $functionName = $button.Tag
                if (Get-Command -Name $functionName -ErrorAction SilentlyContinue) {
                    $button.Add_Click({
                            & $functionName
                        })
                }
                else {
                    Write-Host "No function found for tag: $($button.Tag)"
                }
            }

            $childWindow.ShowDialog()
        }
        catch {
            Write-Host "Failed to load child window XAML: $_" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Error: File not found - $xamlPath" -ForegroundColor Red
    }
}