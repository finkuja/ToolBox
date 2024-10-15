function Remove-Office {
    # Inform the user that the removal process is starting
    $result = [System.Windows.Forms.MessageBox]::Show(
        "This action will close all running Office apps and remove all Office instances. Please save any important work before proceeding. Do you want to continue?", 
        "Confirmation", 
        [System.Windows.Forms.MessageBoxButtons]::YesNo, 
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host "Removing Office..." -ForegroundColor Green

        # Arrays to store paths that couldn't be deleted
        $failedFolders = @()
        $failedRegistryPaths = @()
        $failedShortcuts = @()
        $failedTempFiles = @()

        # Step 1: Uninstall Office using the Office Removal Tool
        try {
            $officeRemovalToolPath = "$env:TEMP\OfficeRemovalTool.exe"
            Write-Host "Downloading Office Removal Tool..." -ForegroundColor Green
            Invoke-WebRequest -Uri "https://aka.ms/SaRA-officeUninstallFromPC" -OutFile $officeRemovalToolPath
            Write-Host "Running Office Removal Tool..." -ForegroundColor Green
            Start-Process -FilePath $officeRemovalToolPath -ArgumentList "/quiet" -Wait
            Write-Host "Office has been uninstalled." -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to uninstall Office using the Office Removal Tool." -ForegroundColor Red
        }

        # Step 2: Remove Office-related folders
        $officeFolders = @(
            "$env:ProgramFiles\Microsoft Office",
            "$env:ProgramFiles (x86)\Microsoft Office",
            "$env:ProgramData\Microsoft\Office",
            "$env:LOCALAPPDATA\Microsoft\Office",
            "$env:APPDATA\Microsoft\Office"
        )
        foreach ($folder in $officeFolders) {
            if (Test-Path $folder) {
                try {
                    Write-Host "Removing folder: $folder" -ForegroundColor Green
                    Remove-Item -Recurse -Force -Path $folder -ErrorAction Stop
                    Write-Host "Removed folder: $folder" -ForegroundColor Green
                }
                catch {
                    Write-Host "Failed to remove folder: $folder" -ForegroundColor Red
                    $failedFolders += $folder
                }
            }
        }

        # Step 3: Remove Office-related registry entries
        $officeRegistryPaths = @(
            "HKCU:\Software\Microsoft\Office",
            "HKCU:\Software\Microsoft\Office\16.0",
            "HKCU:\Software\Microsoft\Office\15.0",
            "HKCU:\Software\Microsoft\Office\14.0",
            "HKCU:\Software\Microsoft\Office\13.0",
            "HKCU:\Software\Microsoft\Office\12.0",
            "HKCU:\Software\Microsoft\Office\11.0",
            "HKLM:\Software\Microsoft\Office",
            "HKLM:\Software\Wow6432Node\Microsoft\Office"
        )
        foreach ($regPath in $officeRegistryPaths) {
            if (Test-Path $regPath) {
                try {
                    Write-Host "Removing registry path: $regPath" -ForegroundColor Green
                    Remove-Item -Recurse -Force -Path $regPath -ErrorAction Stop
                    Write-Host "Removed registry path: $regPath" -ForegroundColor Green
                }
                catch {
                    Write-Host "Failed to remove registry path: $regPath" -ForegroundColor Red
                    $failedRegistryPaths += $regPath
                }
            }
        }

        # Step 4: Remove Office shortcuts
        $officeShortcuts = @(
            "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Office",
            "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office"
        )
        foreach ($shortcut in $officeShortcuts) {
            if (Test-Path $shortcut) {
                try {
                    Write-Host "Removing shortcut: $shortcut" -ForegroundColor Green
                    Remove-Item -Recurse -Force -Path $shortcut -ErrorAction Stop
                    Write-Host "Removed shortcut: $shortcut" -ForegroundColor Green
                }
                catch {
                    Write-Host "Failed to remove shortcut: $shortcut" -ForegroundColor Red
                    $failedShortcuts += $shortcut
                }
            }
        }

        # Step 5: Remove Office temp files and cache files
        $officeTempFiles = @(
            "$env:TEMP\*Office*",
            "$env:TEMP\*MSO*",
            "$env:LOCALAPPDATA\Temp\*Office*",
            "$env:LOCALAPPDATA\Temp\*MSO*"
        )
        foreach ($tempFile in $officeTempFiles) {
            if (Test-Path $tempFile) {
                try {
                    Write-Host "Removing temp file: $tempFile" -ForegroundColor Green
                    Remove-Item -Recurse -Force -Path $tempFile -ErrorAction Stop
                    Write-Host "Removed temp file: $tempFile" -ForegroundColor Green
                }
                catch {
                    Write-Host "Failed to remove temp file: $tempFile" -ForegroundColor Red
                    $failedTempFiles += $tempFile
                }
            }
        }

        # Display summary of failed deletions
        if ($failedFolders.Count -gt 0) {
            Write-Host "The following folders could not be deleted:" -ForegroundColor Yellow
            $failedFolders | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
        }

        if ($failedRegistryPaths.Count -gt 0) {
            Write-Host "The following registry paths could not be deleted:" -ForegroundColor Yellow
            $failedRegistryPaths | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
        }

        if ($failedShortcuts.Count -gt 0) {
            Write-Host "The following shortcuts could not be deleted:" -ForegroundColor Yellow
            $failedShortcuts | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
        }

        if ($failedTempFiles.Count -gt 0) {
            Write-Host "The following temp files could not be deleted:" -ForegroundColor Yellow
            $failedTempFiles | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
        }

        Write-Host "Office removal process is complete." -ForegroundColor Green

        # Prompt the user to reboot the computer
        [System.Windows.Forms.MessageBox]::Show(
            "The Office removal process is complete. Please reboot your computer to finalize the changes.", 
            "Reboot Required", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
}