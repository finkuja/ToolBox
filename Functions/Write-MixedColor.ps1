<#
.SYNOPSIS
Writes text with mixed rainbow colors to the console.

.DESCRIPTION
The Write-MixedColor function writes text to the console with each character in a different color, creating a rainbow effect. 
It cycles through a predefined set of colors to achieve the mixed color effect.

.PARAMETER Text
Specifies the text to be written to the console.

.NOTES
File Name      : Write-MixedColor.ps1
The function uses the Write-Host cmdlet to write colored text to the console.
The function is part of the ToolBox project and is stored in the GitHub repository https://github.com/finkuja/ToolBox

.EXAMPLE
Write-MixedColor -Text "Hello, World!"
Writes "Hello, World!" to the console with each character in a different color.

#>
function Write-MixedColorTitle {
    param (
        [string]$Text
    )
    $colors = @("Red", "Yellow", "Green", "Blue", "Cyan", "Magenta")
    $colorIndex = 0
    $fixedColor = "Gray"
    $rainbowChars = @('|', '\', '/', '_', '-', '(', ')')

    foreach ($char in $Text.ToCharArray()) {
        if ($rainbowChars -contains $char) {
            Write-Host -NoNewline $char -ForegroundColor $colors[$colorIndex]
            $colorIndex = ($colorIndex + 1) % $colors.Length
        }
        else {
            Write-Host -NoNewline $char -ForegroundColor $fixedColor
        }
    }
    Write-Host ""
}

# Function to write text with mixed colors for subtitles
function Write-MixedColorSubtitle {
    param (
        [string]$Text
    )
    $colors = @("Red", "Yellow", "Green", "Blue", "Cyan", "Magenta")
    $colorIndex = 0
    $fixedColor = "DarkYellow"
    $rainbowChars = @('=', '+')

    foreach ($char in $Text.ToCharArray()) {
        if ($rainbowChars -contains $char) {
            Write-Host -NoNewline $char -ForegroundColor $colors[$colorIndex]
            $colorIndex = ($colorIndex + 1) % $colors.Length
        }
        else {
            Write-Host -NoNewline $char -ForegroundColor $fixedColor
        }
    }
    Write-Host ""
}