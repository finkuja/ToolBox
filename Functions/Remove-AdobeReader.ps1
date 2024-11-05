# Function to fully remove all Adobe Reader instances
#----------------------------------------------------
<# .SYNOPSIS
    Fully removes all instances of Adobe Reader from the system.

.DESCRIPTION
    The Remove-AdobeReader function performs a thorough removal of all Adobe Reader instances from the system. 
    It uses the Adobe Reader removal tool to uninstall all versions, followed by using Get-WinGetPackage to uninstall any remaining instances.
    Additionally, it removes related folders and registry entries to ensure a clean uninstallation.

.EXAMPLE
    Remove-AdobeReader
    This command will remove all instances of Adobe Reader from the system, including related folders and registry entries.

.NOTES
    This function requires an internet connection to download the Adobe Reader removal tool if it is not already present in the TEMP directory.
    It also requires the Microsoft.WinGet.Client module to be installed and imported.

.OUTPUTS
    Outputs the status of the operation to the console.
#>
function Remove-AdobeReader {

    # Create a confirmation dialog
    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to remove Adobe Reader?",
        "Confirmation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirmation -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host "Operation cancelled by the user."
        return
    }

    Write-Host "Removing Adobe Reader..."
        
    # Step 1: Uninstall Adobe Reader using the Adobe AcroCleaner tool
    try {
        $acroCleanerToolPath = "$env:TEMP\AcroCleaner_DC2021.exe"
        if (-not (Test-Path $acroCleanerToolPath)) {
            Write-Host "Downloading the Adobe AcroCleaner tool..." -ForegroundColor Yellow
            Invoke-WebRequest -Uri "https://ardownload2.adobe.com/pub/adobe/acrobat/win/AcrobatDC/2100120135/x64/AdobeAcroCleaner_DC2021.exe" -OutFile $acroCleanerToolPath
        }
        Write-Host "Running the Adobe AcroCleaner tool..." -ForegroundColor Yellow
        Start-Process -FilePath $acroCleanerToolPath -ArgumentList "/silent" -NoNewWindow -Wait
        Write-Host "Adobe AcroCleaner tool has completed." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to run the Adobe AcroCleaner tool. Exiting function." -ForegroundColor Red
        return
    }

    # Step 2: Uninstall Adobe Reader using Get-WinGetPackage
    try {
        # Get all installed Adobe Reader packages
        $adobePackages = Get-WinGetPackage | Where-Object { $_.Name -like "*Adobe Acrobat Reader*" }
        if ($adobePackages) {
            foreach ($package in $adobePackages) {
                $packageId = $package.Id
                Write-Host "Uninstalling Adobe Reader package: $packageId"
                Start-Process -FilePath "winget" -ArgumentList "uninstall --id $packageId -e --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait
            }
            Write-Host "All versions of Adobe Reader have been uninstalled using winget." -ForegroundColor Green
        }
        else {
            Write-Host "No Adobe Reader packages found to uninstall using winget." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Failed to uninstall Adobe Reader using winget. Exiting function." -ForegroundColor Red
        return
    }

    # Step 3: Remove Adobe Reader-related folders
    $adobeReaderFolders = @(
        "$env:ProgramFiles\Adobe\Acrobat Reader DC",
        "$env:ProgramFiles (x86)\Adobe\Acrobat Reader DC",
        "$env:ProgramData\Adobe\Acrobat",
        "$env:LOCALAPPDATA\Adobe\Acrobat",
        "$env:APPDATA\Adobe\Acrobat"
    )
    $lockedFolders = @()
    foreach ($folder in $adobeReaderFolders) {
        if (Test-Path $folder) {
            try {
                Remove-Item -Recurse -Force -Path $folder
                Write-Host "Removed folder: $folder" -ForegroundColor Green
            }
            catch {
                Write-Host "Failed to remove folder: $folder" -ForegroundColor Red
                $lockedFolders += $folder
            }
        }
        else {
            Write-Host "Path not found: $folder" -ForegroundColor Yellow
        }
    }

    if ($lockedFolders.Count -gt 0) {
        Write-Host "The following folders were not deleted because they are in use:" -ForegroundColor Yellow
        $lockedFolders | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
    }
    else {
        Write-Host "All Adobe Reader-related folders were successfully deleted." -ForegroundColor Green
    }

    Write-Host "Adobe Reader uninstallation process completed." -ForegroundColor Green
}