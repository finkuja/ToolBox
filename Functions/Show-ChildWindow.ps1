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

        # Find the close button and add a click event handler
        $closeButton = $childWindow.FindName("CloseButton")
        if ($null -ne $closeButton) {
            $closeButton.Add_Click({
                    $childWindow.Close()
                })
        }

        # Load button-function mappings from JSON file
        $jsonContent = Get-Content -Path $jsonPath -Raw | ConvertFrom-Json

        # Initialize hashtables for controls and functions
        $controlHashtable = @{}
        $functionHashtable = @{}

        foreach ($index in 0..($jsonContent.Count - 1)) {
            $mapping = $jsonContent[$index]
            $buttonName = $mapping.ButtonName
            $functionName = $mapping.FunctionName

            $control = $childWindow.FindName($buttonName)
            if ($null -eq $control) {
                #Write-Host "$buttonName not found in XAML." -ForegroundColor Red
                continue
            }

            # Store control and function in hashtables using the index as the key
            $controlHashtable[$index] = $control
            $functionHashtable[$index] = $functionName

            # Add click event handler
            $control.Add_Click({
                    param ($localSender, $e)
                    $currentIndex = $localSender.Tag
                    $currentFunctionName = $functionHashtable[$currentIndex]
                    #Write-Host "Executing function: $currentFunctionName"
                    & $currentFunctionName
                })

            # Set the Tag property to the current index
            $control.Tag = $index
            
        }

        # Show the child window
        $childWindow.ShowDialog()
    }
    catch {
        Write-Host "Failed to load child window XAML: $_" -ForegroundColor Red
    }
}