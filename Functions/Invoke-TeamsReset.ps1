<#
.SYNOPSIS
Resets Microsoft Teams.

.DESCRIPTION
The Invoke-TeamsReset function resets the Microsoft Teams application. 
It prompts the user for confirmation before proceeding, checks if Teams is running, terminates it if it is, 
resets the Teams app, and restarts it if it was running before.

.PARAMETER None
This function does not take any parameters.

.NOTES
File Name      : Invoke-TeamsReset.ps1
The function uses the Start-Process cmdlet to run the reset tasks in an elevated PowerShell session.
The function is part of the ToolBox project and is stored in the GitHub repository https://github.com/finkuja/ToolBox

.EXAMPLE
Invoke-TeamsReset
Prompts the user for confirmation and runs the Teams reset process.

#>
function Invoke-TeamsReset {
    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to reset Microsoft Teams?",
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

    # Reset Teams app
    try {
        Get-AppxPackage -Name "MicrosoftTeams" | Reset-AppxPackage
        Write-Host "Teams app has been reset." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to reset Teams app. Error: $_" -ForegroundColor Red
    }

    # Restart Teams if it was running before
    foreach ($teamProcess in $teamsRunning) {
        Start-Process $teamProcess.Name
        Write-Host "$($teamProcess.Name) has been restarted." -ForegroundColor Green
    }

    [System.Windows.Forms.MessageBox]::Show("The process is complete. Outlook and Teams have been restarted if they were running before.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}
