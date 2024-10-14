<#
.SYNOPSIS
Displays a child window defined by a XAML file.

.PARAMETER xamlPath
The file path to the XAML file that defines the child window.

.DESCRIPTION
The Show-ChildWindow function loads a XAML file, creates a child window, and assigns event handlers to buttons within the window. 
It looks for a button named "CloseButton" and assigns a click event handler to close the window. 
It also assigns click event handlers to other buttons based on their tags, which should correspond to existing functions.

.EXAMPLE
Show-ChildWindow -xamlPath "C:\Path\To\Your\ChildWindow.xaml"
This example loads and displays the child window defined in the specified XAML file.

.NOTES
Ensure that the XAML file exists at the specified path and that the buttons have appropriate tags corresponding to existing functions.

#>
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