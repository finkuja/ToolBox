<#
.SYNOPSIS
Removes all Windows Hello for Business (WHfB) containers.

.DESCRIPTION
The Remove-WHfBContainer function removes all Windows Hello for Business (WHfB) containers from the system. 
It uses the certutil command to delete the Hello containers.

.PARAMETER None
This function does not take any parameters.

.NOTES
File Name      : Remove-WHfBContainer.ps1
The function uses the certutil command to remove WHfB containers.
The function is part of the ToolBox project and is stored in the GitHub repository https://github.com/finkuja/ToolBox

.EXAMPLE
Remove-HelloContainer
Prompts the user and removes all Windows Hello for Business (WHfB) containers from the system.

#>
function Remove-HelloContainer {
    Add-Type -AssemblyName System.Windows.Forms

    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to remove all Windows Hello for Business (WHfB) containers?",
        "Confirmation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirmation -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host "Operation cancelled by the user."
        return
    }
    # Inform the user that the removal process is starting
    Write-Host "Removing all Windows Hello for Business (WHfB) containers..." -ForegroundColor Yellow

    try {
        # Execute the certutil command to delete Hello containers
        certutil -deletehellocontainer

        Write-Host "All Windows Hello for Business (WHfB) containers have been removed." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to remove Windows Hello for Business (WHfB) containers. Error: $_" -ForegroundColor Red
    }
}
