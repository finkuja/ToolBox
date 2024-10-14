function Invoke-DiskCleanUp {
    # Run disk cleanup
    Start-Process "cleanmgr.exe" -ArgumentList "/sagerun:1" -NoNewWindow -Wait

    # Run DISM command with /StartComponentCleanup
    Start-Process "dism.exe" -ArgumentList "/Online /Cleanup-Image /StartComponentCleanup" -NoNewWindow -Wait
}