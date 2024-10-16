<#
.SYNOPSIS
Performs a comprehensive system repair.

.DESCRIPTION
The Invoke-SystemRepair function performs a series of system repair tasks, including CHKDSK, SFC scans, and DISM. 
It prompts the user for confirmation before proceeding and runs the repair tasks with elevated privileges.

.PARAMETER None
This function does not take any parameters.

.NOTES
File Name      : Invoke-SystemRepair.ps1
The function uses the Start-Process cmdlet to run the repair tasks in an elevated PowerShell session.
The function is part of the ToolBox project and is stored in the GitHub repository https://github.com/finkuja/ToolBox

.EXAMPLE
Invoke-SystemRepair
Prompts the user for confirmation and runs the system repair tasks.

#>
function Invoke-SystemRepair {
    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to perform a system repair? This may take some time.",
        "Confirmation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirmation -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host "Operation cancelled by the user."
        return
    }
    Start-Process PowerShell -ArgumentList {
        Write-Host '(1/4) Chkdsk' -ForegroundColor Green
        Chkdsk /scan
        Write-Host '`n(2/4) SFC - 1st scan' -ForegroundColor Green
        sfc /scannow
        Write-Host '`n(3/4) DISM' -ForegroundColor Green
        DISM /Online /Cleanup-Image /Restorehealth
        Write-Host '`n(4/4) SFC - 2nd scan' -ForegroundColor Green
        sfc /scannow
        Read-Host '`nPress Enter to Continue'
    } -Verb RunAs
}