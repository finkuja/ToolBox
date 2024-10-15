function Invoke-TeamsAddinFix {
    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to proceed with the Teams add-in fix?",
        "Confirmation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirmation -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host "Operation cancelled by the user."
        return
    }

    # Check if Teams and Outlook are running
    $teamsRunning = @()
    $teamsRunning += Get-Process -Name Teams -ErrorAction SilentlyContinue
    $teamsRunning += Get-Process -Name ms-teams -ErrorAction SilentlyContinue
    $teamsRunning = $teamsRunning | Where-Object { $_ -ne $null }
    $outlookRunning = Get-Process -Name Outlook -ErrorAction SilentlyContinue

    # Terminate Teams and Outlook
    foreach ($teamProcess in $teamsRunning) {
        Stop-Process -Id $teamProcess.Id -Force
        Write-Host "$($teamProcess.Name) has been terminated." -ForegroundColor Green
    }

    if ($outlookRunning) {
        Stop-Process -Name Outlook -Force
        Write-Host "Outlook has been terminated." -ForegroundColor Green
    }
    else {
        Write-Host "Outlook was not running." -ForegroundColor Yellow
    }

    # Step 1: Remove SquirrelTemp and Teams folders
    try {
        Remove-Item -Recurse -Force -Path "$env:LOCALAPPDATA\SquirrelTemp"
        Write-Host "SquirrelTemp folder has been removed." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to remove SquirrelTemp folder." -ForegroundColor Red
    }

    try {
        Remove-Item -Recurse -Force -Path "$env:LOCALAPPDATA\Microsoft\Teams"
        Write-Host "Teams folder has been removed." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to remove Teams folder." -ForegroundColor Red
    }

    # Step 2: Rename tma_settings.json
    $tmaSettingsPath = "$env:LOCALAPPDATA\Publishers\8wekyb3d8bbwe\TeamsSharedConfig\tma_settings.json"
    if (Test-Path $tmaSettingsPath) {
        try {
            Rename-Item -Path $tmaSettingsPath -NewName "tma_settings.json.old"
            Write-Host "tma_settings.json has been renamed." -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to rename tma_settings.json." -ForegroundColor Red
        }
    }
    else {
        Write-Host "tma_settings.json not found." -ForegroundColor Yellow
    }

    # Step 3: Re-register Microsoft.Teams.AddinLoader.dll
    try {
        if ([Environment]::Is64BitOperatingSystem) {
            & "$env:SystemRoot\System32\regsvr32.exe" /n /i:user "$env:LOCALAPPDATA\Microsoft\TeamsMeetingAddin\1.0.18012.2\x64\Microsoft.Teams.AddinLoader.dll"
        }
        else {
            & "$env:SystemRoot\SysWOW64\regsvr32.exe" /n /i:user "$env:LOCALAPPDATA\Microsoft\TeamsMeetingAddin\1.0.18012.2\x86\Microsoft.Teams.AddinLoader.dll"
        }
        Write-Host "Microsoft.Teams.AddinLoader.dll has been re-registered." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to re-register Microsoft.Teams.AddinLoader.dll." -ForegroundColor Red
    }

    # Step 4: Check and set LoadBehavior in the registry
    $regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect"
    if (Test-Path $regPath) {
        try {
            $loadBehavior = Get-ItemProperty -Path $regPath -Name LoadBehavior
            if ($loadBehavior.LoadBehavior -ne 3) {
                Set-ItemProperty -Path $regPath -Name LoadBehavior -Value 3
            }
            Write-Host "LoadBehavior has been set to 3." -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to set LoadBehavior." -ForegroundColor Red
        }
    }
    else {
        Write-Host "Registry path for LoadBehavior not found." -ForegroundColor Yellow
    }

    # Step 5: Reset Teams UWP app
    try {
        Get-AppxPackage -Name "MicrosoftTeams" | Reset-AppxPackage
        Write-Host "Teams UWP app has been reset." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to reset Teams UWP app." -ForegroundColor Red
    }

    # Reopen Teams and Outlook if they were running before
    foreach ($teamProcess in $teamsRunning) {
        Start-Process $teamProcess.Name
        Write-Host "$($teamProcess.Name) has been restarted." -ForegroundColor Green
    }

    if ($outlookRunning) {
        Start-Process "Outlook"
        Write-Host "Outlook has been restarted." -ForegroundColor Green
    }

    [System.Windows.Forms.MessageBox]::Show("The process is complete. Teams and Outlook have been restarted if they were running before.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}