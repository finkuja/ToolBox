<#
.SYNOPSIS
Displays a child window defined by a XAML file.

.PARAMETER xamlPath
The file path to the XAML file that defines the child window.

.PARAMETER jsonPath
The file path to the JSON file that defines the button-function mappings.

.DESCRIPTION
The Show-ChildWindow function loads a XAML file, creates a child window, and assigns event handlers to buttons within the window. 
It looks for a button named "CloseButton" and assigns a click event handler to close the window. 
It also assigns click event handlers to other buttons based on their tags, which should correspond to existing functions.

.EXAMPLE
Show-ChildWindow -xamlPath "C:\Path\To\Your\ChildWindow.xaml" -jsonPath "C:\Path\To\Your\ButtonMappings.json"
This example loads and displays the child window defined in the specified XAML file and uses the specified JSON file for button-function mappings.

.NOTES
Ensure that the XAML file exists at the specified path and that the buttons have appropriate tags corresponding to existing functions.

#>
function Show-ChildWindow {
    param (
        [string]$xamlPath,
        [string]$jsonPath
    )

    if (-not (Test-Path $xamlPath)) {
        Write-Host "Error: XAML file not found - $xamlPath" -ForegroundColor Red
        return
    }

    if (-not (Test-Path $jsonPath)) {
        Write-Host "Error: JSON file not found - $jsonPath" -ForegroundColor Red
        return
    }

    try {
        $childXaml = Get-Content -Path $xamlPath -Raw
        $childReader = New-Object System.Xml.XmlTextReader (New-Object System.IO.StringReader $childXaml)
        $childWindow = [Windows.Markup.XamlReader]::Load($childReader)
        
        if ($null -eq $childWindow) {
            throw "Failed to load child window from XAML."
        }

        # Load button-function mappings from JSON file
        $jsonContent = Get-Content -Path $jsonPath -Raw | ConvertFrom-Json
        $childButtons = @{}
        foreach ($mapping in $jsonContent) {
            $childButtons[$mapping.ButtonName] = $mapping.FunctionName
        }

        # Find the close button and add a click event handler
        $closeButton = $childWindow.FindName("CloseButton")
        if ($null -ne $closeButton) {
            $closeButton.Add_Click({
                    $childWindow.Close()
                })
        }

        # Find all buttons in the child window based on the hashtable keys
        foreach ($name in $childButtons.Keys) {
            $button = $childWindow.FindName($name)
            if ($button -is [System.Windows.Controls.Button]) {
                $functionName = $childButtons[$name]
                if (Get-Command -Name $functionName -ErrorAction SilentlyContinue) {
                    $button.Add_Click({
                            & $functionName
                        })
                }
                else {
                    Write-Host "No function found for button: $name"
                }
            }
            else {
                Write-Host "No button found with name: $name"
            }
        }

        $childWindow.ShowDialog()
    }
    catch {
        Write-Host "Failed to load child window XAML: $_" -ForegroundColor Red
    }
}