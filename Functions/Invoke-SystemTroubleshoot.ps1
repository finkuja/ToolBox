function Invoke-SystemTroubleshoot {
    # Open Windows Other Troubleshooters settings page
    Write-Host "Opening Windows Other Troubleshooters..." -ForegroundColor Yellow
    Start-Process ms-settings:troubleshoot-other
}