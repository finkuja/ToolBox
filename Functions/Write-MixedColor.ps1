# Mixed color functions for writing text with rainbow colors
function Write-MixedColorTitle {
    param (
        [string]$Text
    )
    $colors = @("Red", "Yellow", "Green", "Blue", "Cyan", "Magenta")
    $colorIndex = 0
    $fixedColor = "DarkBlue"
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