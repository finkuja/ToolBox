<#
.SYNOPSIS
    This script performs disk cleanup tasks.

.DESCRIPTION
    The script executes a series of commands to clean up disk space by removing unnecessary files and optimizing storage. It can be used to automate the process of freeing up disk space and improving system performance.

.PARAMETER <None>
    This function does not take any parameters.

.NOTES
    File Name      : Invoke-DiskCleanUp.ps1
    This script is part of the ToolBox project and is stored in the GitHub repository https://github.com/finkuja/ToolBox
.EXAMPLE
    Invoke-DiskCleanUp
    This command runs the disk cleanup tasks, removing unnecessary files and optimizing storage to free up disk space.
#>
function Invoke-DiskCleanUp {
    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to proceed with the disk cleanup? This operation will take some time to complete.",
        "Confirmation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirmation -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host "Operation cancelled by the user."
        return
    }
    # Run disk cleanup
    Start-Process "cleanmgr.exe" -ArgumentList "/sagerun:1" -NoNewWindow -Wait

    # Run DISM command with /StartComponentCleanup
    Start-Process "dism.exe" -ArgumentList "/Online /Cleanup-Image /StartComponentCleanup" -NoNewWindow -Wait
}