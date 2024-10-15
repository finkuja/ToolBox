function Invoke-TeamsCacheClean {
    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to clean the Teams cache?",
        "Confirmation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirmation -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host "Operation cancelled by the user."
        return
    }

    # Check if Teams and ms-teams are running
    $teamsRunning = @()
    $teamsRunning += Get-Process -Name Teams -ErrorAction SilentlyContinue
    $teamsRunning += Get-Process -Name ms-teams -ErrorAction SilentlyContinue
    $teamsRunning = $teamsRunning | Where-Object { $_ -ne $null }

    # Terminate Teams if running
    foreach ($teamProcess in $teamsRunning) {
        Stop-Process -Id $teamProcess.Id -Force
        Write-Host "$($teamProcess.Name) has been terminated." -ForegroundColor Green
    }

    # Clean Teams cache
    $teamsCachePath = "$env:APPDATA\Microsoft\Teams"
    if (Test-Path $teamsCachePath) {
        Remove-Item -Recurse -Force -Path "$teamsCachePath\*"
        Write-Host "Teams cache has been cleaned." -ForegroundColor Green
    }
    else {
        Write-Host "Teams cache path not found." -ForegroundColor Yellow
    }

    # Restart Teams if it was running before
    foreach ($teamProcess in $teamsRunning) {
        Start-Process $teamProcess.Name
        Write-Host "$($teamProcess.Name) has been restarted." -ForegroundColor Green
    }

    [System.Windows.Forms.MessageBox]::Show("The process is complete. Outlook and Teams have been restarted if they were running before.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}
