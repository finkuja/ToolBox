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