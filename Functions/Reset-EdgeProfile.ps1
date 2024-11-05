<#
.SYNOPSIS
Resets the Microsoft Edge browser profile.

.DESCRIPTION
The Reset-EdgeProfile function resets the user profile for Microsoft Edge. 
It prompts the user for confirmation before proceeding, and then resets the Edge profile using the `--reset-profile` argument.

.PARAMETER None
This function does not take any parameters.

.NOTES
File Name      : Reset-EdgeProfile.ps1
The function uses the Start-Process cmdlet to run the Edge browser with the `--reset-profile` argument.
The function is part of the ToolBox project and is stored in the GitHub repository https://github.com/finkuja/ToolBox

.EXAMPLE
Reset-EdgeProfile
Prompts the user for confirmation and resets the Microsoft Edge browser profile.

#>
function Reset-EdgeProfile {
    # Create a confirmation dialog
    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to reset the Microsoft Edge browser profile?",
        "Confirmation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirmation -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host "Operation cancelled by the user."
        return
    }
    try {
        # Step 1: Reset Edge profile silently
        Write-Host "Resetting Edge browser profile..."
        Start-Process "msedge" -ArgumentList "--reset-profile" -NoNewWindow -Wait
        Write-Host "Edge browser profile has been reset." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to reset Edge profile. Error: $_" -ForegroundColor Red
    }
}