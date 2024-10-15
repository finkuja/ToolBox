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