function Invoke-TeamsRemoval {
    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to uninstall Microsoft Teams and delete all related files and registry keys?",
        "Confirmation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirmation -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host "Operation cancelled by the user."
        return
    }
    Write-Host "Stopping Teams processes..."
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

    Write-Host "Removing Teams..."

    # Uninstall Teams
    $teamsUninstallPath = "$env:LOCALAPPDATA\Microsoft\Teams\Update.exe"
    if (Test-Path $teamsUninstallPath) {
        Start-Process $teamsUninstallPath -ArgumentList "--uninstall" -Wait
        Write-Host "Teams has been removed." -ForegroundColor Green
    }
    else {
        Write-Host "Teams uninstall path not found." -ForegroundColor Yellow
    }

    # Delete Teams cache and temp files
    $teamsCachePath = "$env:APPDATA\Microsoft\Teams"
    if (Test-Path $teamsCachePath) {
        Remove-Item -Recurse -Force -Path "$teamsCachePath\*"
        Write-Host "Teams cache has been deleted." -ForegroundColor Green
    }
    else {
        Write-Host "Teams cache path not found." -ForegroundColor Yellow
    }

    $teamsTempPath = "$env:LOCALAPPDATA\Microsoft\Teams"
    if (Test-Path $teamsTempPath) {
        Remove-Item -Recurse -Force -Path "$teamsTempPath\*"
        Write-Host "Teams temp files have been deleted." -ForegroundColor Green
    }
    else {
        Write-Host "Teams temp path not found." -ForegroundColor Yellow
    }

    # Delete Teams registry keys
    $teamsRegKeys = @(
        "HKCU:\Software\Microsoft\Office\Teams",
        "HKCU:\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect",
        "HKCU:\Software\Microsoft\Teams"
    )
    foreach ($regKey in $teamsRegKeys) {
        if (Test-Path $regKey) {
            Remove-Item -Path $regKey -Recurse -Force
            Write-Host "Registry key $regKey has been deleted." -ForegroundColor Green
        }
        else {
            Write-Host "Registry key $regKey not found." -ForegroundColor Yellow
        }
    }

    # Inform the user of successful completion and suggest rebooting
    $rebootResult = [System.Windows.Forms.MessageBox]::Show("Teams has been successfully uninstalled and related files and registry keys have been deleted. It is recommended to reboot your system. Do you want to reboot now?", "Information", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information)

    if ($rebootResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        Restart-Computer -Force
    }
    else {
        Write-Host "Reboot operation was canceled by the user." -ForegroundColor Yellow
    }
    [System.Windows.Forms.MessageBox]::Show("The process is complete. Outlook and Teams have been restarted if they were running before.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}
