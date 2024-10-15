function Reset-Network {
    # Reset network
    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to reset the network settings?",
        "Confirmation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirmation -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host "Network reset operation cancelled by the user."
        return
    }

    Write-Host "Resetting network..." -ForegroundColor Yellow

    # Reset network using netsh
    netsh winsock reset | Out-Null
    netsh int ip reset | Out-Null
    netsh interface ipv4 reset | Out-Null
    netsh interface ipv6 reset | Out-Null
    netsh interface reset all | Out-Null

    # Reset all network adapters
    $networkAdapters = Get-NetAdapter
    foreach ($adapter in $networkAdapters) {
        $adapter | Disable-NetAdapter -Confirm:$false | Out-Null
        $adapter | Enable-NetAdapter -Confirm:$false | Out-Null
    }

    $rebootPrompt = [System.Windows.Forms.MessageBox]::Show(
        "Network reset is complete. It is recommended to reboot your computer to apply all changes. Do you want to reboot now?", 
        "Reboot Required", 
        [System.Windows.Forms.MessageBoxButtons]::YesNo, 
        [System.Windows.Forms.MessageBoxIcon]::Information
    )

    if ($rebootPrompt -eq [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host "Rebooting the system..." -ForegroundColor Green
        Restart-Computer -Force
    }
    else {
        Write-Host "Please reboot your computer manually to apply all changes." -ForegroundColor Yellow
    }
}