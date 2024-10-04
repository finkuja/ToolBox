<#
.SYNOPSIS
    This script is the ESSToolBox, a tool for system administration tasks.

.DESCRIPTION
    The ESSToolBox is a PowerShell script designed to perform various system administration tasks. It includes features such as repairing Windows Update, checking for the installation of Windows Package Manager (winget), and more.

.PARAMETER None
    This script does not accept any parameters.

.NOTES
    - This script must be run with administrator privileges.
    - The script checks if Windows Package Manager (winget) is installed and prompts the user to install it if necessary.
    - The script also checks for the 'msstore' source in the winget source list and accepts the source agreements if necessary.

.LINK
    GitHub Repository: https://github.com/finkuja/ToolBox
#>
# Ensure you run this script with administrator privileges


########################################################
# ESSToolBox - A PowerShell System Administration Tool #
########################################################

Write-Output " ______  _____  _____    _______          _   ____ 
|  ____|/ ____|/ ____|  |__   __|        | | |  _ \             
| |__  | (___ | (___       | | ___   ___ | | | |_) | ___  __  __
|  __|  \___ \ \___ \      | |/ _ \ / _ \| | |  _ < / _ \ \ \/ /
| |____ ____) |____) |     | | (_) | (_) | | | |_) | (_) | |  |
|______|_____/|_____/      |_|\___/ \___/|_| |____/ \___/ /_/\_\
 
 === Version Beta 0.1 ===

 === Created by: Carlos Alvarez Magariños ===
 "
# Check if the script is running with administrator privileges
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    Write-Host "The ESSToolBox needs to run as Administrator, trying to elevate the permissions..." -ForegroundColor Yellow
    
    # Get the current script content
    $scriptContent = (irm https://raw.githubusercontent.com/finkuja/ToolBox/refs/heads/main/EssTool.ps1)
    
    # Define a temporary file path
    $tempFilePath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "ESSToolBox.ps1")
    
    # Save the script content to the temporary file
    [System.IO.File]::WriteAllText($tempFilePath, $scriptContent)
    
    # Create a new process to run the script with administrator privileges
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = "powershell.exe"
    $startInfo.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$tempFilePath`""
    $startInfo.Verb = "runas"
    
    try {
        # Start the new process
        $process = [System.Diagnostics.Process]::Start($startInfo)
        $process.WaitForExit()
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to run the script with administrator privileges. Please run this script with administrator privileges manually.")
    }
    
    # Exit the current script
    exit
}
else {
    Write-Host "Running with administrator privileges." -ForegroundColor Green
}

#Check if winget is installed
Write-Host "Checking if Windows Package Manager (winget) is installed..."
$winget = Get-Command winget -ErrorAction SilentlyContinue
if ($null -eq $winget) {
    [System.Windows.Forms.MessageBox]::Show("Windows Package Manager (winget) is not installed. Please install it from https://github.com/microsoft/winget-cli/releases")
    Start-Process "https://github.com/microsoft/winget-cli/releases"
    exit
}
else {
    Write-Host "Windows Package Manager (winget) is installed." -ForegroundColor Green
}

###############################
# PRIVATE FUNCTION DEFINITIONS#
###############################

# Function from Chris Titus Tech winutils.ps1 script github.com/christitus/winutils to Fix Windows Update
#---------------------------------------------------------------------------------------------------------
<#

    .SYNOPSIS
        Performs various tasks in an attempt to repair Windows Update

    .DESCRIPTION
        1. (Aggressive Only) Scans the system for corruption using chkdsk, SFC, and DISM
            Steps:
                1. Runs chkdsk /scan /perf
                    /scan - Runs an online scan on the volume
                    /perf - Uses more system resources to complete a scan as fast as possible
                2. Runs SFC /scannow
                    /scannow - Scans integrity of all protected system files and repairs files with problems when possible
                3. Runs DISM /Online /Cleanup-Image /RestoreHealth
                    /Online - Targets the running operating system
                    /Cleanup-Image - Performs cleanup and recovery operations on the image
                    /RestoreHealth - Scans the image for component store corruption and attempts to repair the corruption using Windows Update
                4. Runs SFC /scannow
                    Ran twice in case DISM repaired SFC
        2. Stops Windows Update Services
        3. Remove the QMGR Data file, which stores BITS jobs
        4. (Aggressive Only) Renames the DataStore and CatRoot2 folders
            DataStore - Contains the Windows Update History and Log Files
            CatRoot2 - Contains the Signatures for Windows Update Packages
        5. Renames the Windows Update Download Folder
        6. Deletes the Windows Update Log
        7. (Aggressive Only) Resets the Security Descriptors on the Windows Update Services
        8. Reregisters the BITS and Windows Update DLLs
        9. Removes the WSUS client settings
        10. Resets WinSock
        11. Gets and deletes all BITS jobs
        12. Sets the startup type of the Windows Update Services then starts them
        13. Forces Windows Update to check for updates

    .PARAMETER Aggressive
        If specified, the script will take additional steps to repair Windows Update that are more dangerous, take a significant amount of time, or are generally unnecessary

    #>
function Invoke-FixesWUpdate {

    param($Aggressive = $false)

    Write-Progress -Id 0 -Activity "Repairing Windows Update" -PercentComplete 0
    # Wait for the first progress bar to show, otherwise the second one won't show
    Start-Sleep -Milliseconds 200

    if ($Aggressive) {
        # Scan system for corruption
        Write-Progress -Id 0 -Activity "Repairing Windows Update" -Status "Scanning for corruption..." -PercentComplete 0
        Write-Progress -Id 1 -ParentId 0 -Activity "Scanning for corruption" -Status "Running chkdsk..." -PercentComplete 0
        # 2>&1 redirects stdout, alowing iteration over the output
        chkdsk.exe /scan /perf 2>&1 | ForEach-Object {
            # Write stdout to the Verbose stream
            Write-Verbose $_

            # Get the index of the total percentage
            $index = $_.IndexOf("Total:")
            if (
                # If the percent is found
                ($percent = try {
                    (
                        $_.Substring(
                            $index + 6,
                            $_.IndexOf("%", $index) - $index - 6
                        )
                    ).Trim()
                }
                catch { 0 }) `
                    <# And the current percentage is greater than the previous one #>`
                    -and $percent -gt $oldpercent
            ) {
                # Update the progress bar
                $oldpercent = $percent
                Write-Progress -Id 1 -ParentId 0 -Activity "Scanning for corruption" -Status "Running chkdsk... ($percent%)" -PercentComplete $percent
            }
        }

        Write-Progress -Id 1 -ParentId 0 -Activity "Scanning for corruption" -Status "Running SFC..." -PercentComplete 0
        $oldpercent = 0
        # SFC has a bug when redirected which causes it to output only when the stdout buffer is full, causing the progress bar to move in chunks
        sfc /scannow 2>&1 | ForEach-Object {
            # Write stdout to the Verbose stream
            Write-Verbose $_

            # Filter for lines that contain a percentage that is greater than the previous one
            if (
                (
                    # Use a different method to get the percentage that accounts for SFC's Unicode output
                    [int]$percent = try {
                        (
                            (
                                $_.Substring(
                                    $_.IndexOf("n") + 2,
                                    $_.IndexOf("%") - $_.IndexOf("n") - 2
                                ).ToCharArray() | Where-Object { $_ }
                            ) -join ''
                        ).TrimStart()
                    }
                    catch { 0 }
                ) -and $percent -gt $oldpercent
            ) {
                # Update the progress bar
                $oldpercent = $percent
                Write-Progress -Id 1 -ParentId 0 -Activity "Scanning for corruption" -Status "Running SFC... ($percent%)" -PercentComplete $percent
            }
        }

        Write-Progress -Id 1 -ParentId 0 -Activity "Scanning for corruption" -Status "Running DISM..." -PercentComplete 0
        $oldpercent = 0
        DISM /Online /Cleanup-Image /RestoreHealth | ForEach-Object {
            # Write stdout to the Verbose stream
            Write-Verbose $_

            # Filter for lines that contain a percentage that is greater than the previous one
            if (
                ($percent = try {
                    [int]($_ -replace "\[" -replace "=" -replace " " -replace "%" -replace "\]")
                }
                catch { 0 }) `
                    -and $percent -gt $oldpercent
            ) {
                # Update the progress bar
                $oldpercent = $percent
                Write-Progress -Id 1 -ParentId 0 -Activity "Scanning for corruption" -Status "Running DISM... ($percent%)" -PercentComplete $percent
            }
        }

        Write-Progress -Id 1 -ParentId 0 -Activity "Scanning for corruption" -Status "Running SFC again..." -PercentComplete 0
        $oldpercent = 0
        sfc /scannow 2>&1 | ForEach-Object {
            # Write stdout to the Verbose stream
            Write-Verbose $_

            # Filter for lines that contain a percentage that is greater than the previous one
            if (
                (
                    [int]$percent = try {
                        (
                            (
                                $_.Substring(
                                    $_.IndexOf("n") + 2,
                                    $_.IndexOf("%") - $_.IndexOf("n") - 2
                                ).ToCharArray() | Where-Object { $_ }
                            ) -join ''
                        ).TrimStart()
                    }
                    catch { 0 }
                ) -and $percent -gt $oldpercent
            ) {
                # Update the progress bar
                $oldpercent = $percent
                Write-Progress -Id 1 -ParentId 0 -Activity "Scanning for corruption" -Status "Running SFC... ($percent%)" -PercentComplete $percent
            }
        }
        Write-Progress -Id 1 -ParentId 0 -Activity "Scanning for corruption" -Status "Completed" -PercentComplete 100
    }


    Write-Progress -Id 0 -Activity "Repairing Windows Update" -Status "Stopping Windows Update Services..." -PercentComplete 10
    # Stop the Windows Update Services
    Write-Progress -Id 2 -ParentId 0 -Activity "Stopping Services" -Status "Stopping BITS..." -PercentComplete 0
    Stop-Service -Name BITS -Force
    Write-Progress -Id 2 -ParentId 0 -Activity "Stopping Services" -Status "Stopping wuauserv..." -PercentComplete 20
    Stop-Service -Name wuauserv -Force
    Write-Progress -Id 2 -ParentId 0 -Activity "Stopping Services" -Status "Stopping appidsvc..." -PercentComplete 40
    Stop-Service -Name appidsvc -Force
    Write-Progress -Id 2 -ParentId 0 -Activity "Stopping Services" -Status "Stopping cryptsvc..." -PercentComplete 60
    Stop-Service -Name cryptsvc -Force
    Write-Progress -Id 2 -ParentId 0 -Activity "Stopping Services" -Status "Completed" -PercentComplete 100


    # Remove the QMGR Data file
    Write-Progress -Id 0 -Activity "Repairing Windows Update" -Status "Renaming/Removing Files..." -PercentComplete 20
    Write-Progress -Id 3 -ParentId 0 -Activity "Renaming/Removing Files" -Status "Removing QMGR Data files..." -PercentComplete 0
    Remove-Item "$env:allusersprofile\Application Data\Microsoft\Network\Downloader\qmgr*.dat" -ErrorAction SilentlyContinue


    if ($Aggressive) {
        # Rename the Windows Update Log and Signature Folders
        Write-Progress -Id 3 -ParentId 0 -Activity "Renaming/Removing Files" -Status "Renaming the Windows Update Log, Download, and Signature Folder..." -PercentComplete 20
        Rename-Item $env:systemroot\SoftwareDistribution\DataStore DataStore.bak -ErrorAction SilentlyContinue
        Rename-Item $env:systemroot\System32\Catroot2 catroot2.bak -ErrorAction SilentlyContinue
    }

    # Rename the Windows Update Download Folder
    Write-Progress -Id 3 -ParentId 0 -Activity "Renaming/Removing Files" -Status "Renaming the Windows Update Download Folder..." -PercentComplete 20
    Rename-Item $env:systemroot\SoftwareDistribution\Download Download.bak -ErrorAction SilentlyContinue

    # Delete the legacy Windows Update Log
    Write-Progress -Id 3 -ParentId 0 -Activity "Renaming/Removing Files" -Status "Removing the old Windows Update log..." -PercentComplete 80
    Remove-Item $env:systemroot\WindowsUpdate.log -ErrorAction SilentlyContinue
    Write-Progress -Id 3 -ParentId 0 -Activity "Renaming/Removing Files" -Status "Completed" -PercentComplete 100


    if ($Aggressive) {
        # Reset the Security Descriptors on the Windows Update Services
        Write-Progress -Id 0 -Activity "Repairing Windows Update" -Status "Resetting the WU Service Security Descriptors..." -PercentComplete 25
        Write-Progress -Id 4 -ParentId 0 -Activity "Resetting the WU Service Security Descriptors" -Status "Resetting the BITS Security Descriptor..." -PercentComplete 0
        Start-Process -NoNewWindow -FilePath "sc.exe" -ArgumentList "sdset", "bits", "D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)"
        Write-Progress -Id 4 -ParentId 0 -Activity "Resetting the WU Service Security Descriptors" -Status "Resetting the wuauserv Security Descriptor..." -PercentComplete 50
        Start-Process -NoNewWindow -FilePath "sc.exe" -ArgumentList "sdset", "wuauserv", "D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)"
        Write-Progress -Id 4 -ParentId 0 -Activity "Resetting the WU Service Security Descriptors" -Status "Completed" -PercentComplete 100
    }


    # Reregister the BITS and Windows Update DLLs
    Write-Progress -Id 0 -Activity "Repairing Windows Update" -Status "Reregistering DLLs..." -PercentComplete 40
    $oldLocation = Get-Location
    Set-Location $env:systemroot\system32
    $i = 0
    $DLLs = @(
        "atl.dll", "urlmon.dll", "mshtml.dll", "shdocvw.dll", "browseui.dll",
        "jscript.dll", "vbscript.dll", "scrrun.dll", "msxml.dll", "msxml3.dll",
        "msxml6.dll", "actxprxy.dll", "softpub.dll", "wintrust.dll", "dssenh.dll",
        "rsaenh.dll", "gpkcsp.dll", "sccbase.dll", "slbcsp.dll", "cryptdlg.dll",
        "oleaut32.dll", "ole32.dll", "shell32.dll", "initpki.dll", "wuapi.dll",
        "wuaueng.dll", "wuaueng1.dll", "wucltui.dll", "wups.dll", "wups2.dll",
        "wuweb.dll", "qmgr.dll", "qmgrprxy.dll", "wucltux.dll", "muweb.dll", "wuwebv.dll"
    )
    foreach ($dll in $DLLs) {
        Write-Progress -Id 5 -ParentId 0 -Activity "Reregistering DLLs" -Status "Registering $dll..." -PercentComplete ($i / $DLLs.Count * 100)
        $i++
        Start-Process -NoNewWindow -FilePath "regsvr32.exe" -ArgumentList "/s", $dll
    }
    Set-Location $oldLocation
    Write-Progress -Id 5 -ParentId 0 -Activity "Reregistering DLLs" -Status "Completed" -PercentComplete 100


    # Remove the WSUS client settings
    if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate") {
        Write-Progress -Id 0 -Activity "Repairing Windows Update" -Status "Removing WSUS client settings..." -PercentComplete 60
        Write-Progress -Id 6 -ParentId 0 -Activity "Removing WSUS client settings" -PercentComplete 0
        Start-Process -NoNewWindow -FilePath "REG" -ArgumentList "DELETE", "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate", "/v", "AccountDomainSid", "/f" -RedirectStandardError $true
        Start-Process -NoNewWindow -FilePath "REG" -ArgumentList "DELETE", "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate", "/v", "PingID", "/f" -RedirectStandardError $true
        Start-Process -NoNewWindow -FilePath "REG" -ArgumentList "DELETE", "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate", "/v", "SusClientId", "/f" -RedirectStandardError $true
        Write-Progress -Id 6 -ParentId 0 -Activity "Removing WSUS client settings" -Status "Completed" -PercentComplete 100
    }


    # Reset WinSock
    Write-Progress -Id 0 -Activity "Repairing Windows Update" -Status "Resetting WinSock..." -PercentComplete 65
    Write-Progress -Id 7 -ParentId 0 -Activity "Resetting WinSock" -Status "Resetting WinSock..." -PercentComplete 0
    Start-Process -NoNewWindow -FilePath "netsh" -ArgumentList "winsock", "reset" -RedirectStandardOutput $true
    Start-Process -NoNewWindow -FilePath "netsh" -ArgumentList "winhttp", "reset", "proxy" -RedirectStandardOutput $true
    Start-Process -NoNewWindow -FilePath "netsh" -ArgumentList "int", "ip", "reset" -RedirectStandardOutput $true
    Write-Progress -Id 7 -ParentId 0 -Activity "Resetting WinSock" -Status "Completed" -PercentComplete 100


    # Get and delete all BITS jobs
    Write-Progress -Id 0 -Activity "Repairing Windows Update" -Status "Deleting BITS jobs..." -PercentComplete 75
    Write-Progress -Id 8 -ParentId 0 -Activity "Deleting BITS jobs" -Status "Deleting BITS jobs..." -PercentComplete 0
    Get-BitsTransfer | Remove-BitsTransfer
    Write-Progress -Id 8 -ParentId 0 -Activity "Deleting BITS jobs" -Status "Completed" -PercentComplete 100


    # Change the startup type of the Windows Update Services and start them
    Write-Progress -Id 0 -Activity "Repairing Windows Update" -Status "Starting Windows Update Services..." -PercentComplete 90
    Write-Progress -Id 9 -ParentId 0 -Activity "Starting Windows Update Services" -Status "Starting BITS..." -PercentComplete 0
    Get-Service BITS | Set-Service -StartupType Manual -PassThru | Start-Service
    Write-Progress -Id 9 -ParentId 0 -Activity "Starting Windows Update Services" -Status "Starting wuauserv..." -PercentComplete 25
    Get-Service wuauserv | Set-Service -StartupType Manual -PassThru | Start-Service
    Write-Progress -Id 9 -ParentId 0 -Activity "Starting Windows Update Services" -Status "Starting AppIDSvc..." -PercentComplete 50
    # The AppIDSvc service is protected, so the startup type has to be changed in the registry
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\AppIDSvc" -Name "Start" -Value "3" # Manual
    Start-Service AppIDSvc
    Write-Progress -Id 9 -ParentId 0 -Activity "Starting Windows Update Services" -Status "Starting CryptSvc..." -PercentComplete 75
    Get-Service CryptSvc | Set-Service -StartupType Manual -PassThru | Start-Service
    Write-Progress -Id 9 -ParentId 0 -Activity "Starting Windows Update Services" -Status "Completed" -PercentComplete 100


    # Force Windows Update to check for updates
    Write-Progress -Id 0 -Activity "Repairing Windows Update" -Status "Forcing discovery..." -PercentComplete 95
    Write-Progress -Id 10 -ParentId 0 -Activity "Forcing discovery" -Status "Forcing discovery..." -PercentComplete 0
    (New-Object -ComObject Microsoft.Update.AutoUpdate).DetectNow()
    Start-Process -NoNewWindow -FilePath "wuauclt" -ArgumentList "/resetauthorization", "/detectnow"
    Write-Progress -Id 10 -ParentId 0 -Activity "Forcing discovery" -Status "Completed" -PercentComplete 100
    Write-Progress -Id 0 -Activity "Repairing Windows Update" -Status "Completed" -PercentComplete 100

    $ButtonType = [System.Windows.MessageBoxButton]::OK
    $MessageboxTitle = "Reset Windows Update "
    $Messageboxbody = ("Stock settings loaded.`n Please reboot your computer")
    $MessageIcon = [System.Windows.MessageBoxImage]::Information

    [System.Windows.MessageBox]::Show($Messageboxbody, $MessageboxTitle, $ButtonType, $MessageIcon)
    Write-Host "==============================================="
    Write-Host "-- Reset All Windows Update Settings to Stock -"
    Write-Host "==============================================="

    # Remove the progress bars
    Write-Progress -Id 0 -Activity "Repairing Windows Update" -Completed
    Write-Progress -Id 1 -Activity "Scanning for corruption" -Completed
    Write-Progress -Id 2 -Activity "Stopping Services" -Completed
    Write-Progress -Id 3 -Activity "Renaming/Removing Files" -Completed
    Write-Progress -Id 4 -Activity "Resetting the WU Service Security Descriptors" -Completed
    Write-Progress -Id 5 -Activity "Reregistering DLLs" -Completed
    Write-Progress -Id 6 -Activity "Removing WSUS client settings" -Completed
    Write-Progress -Id 7 -Activity "Resetting WinSock" -Completed
    Write-Progress -Id 8 -Activity "Deleting BITS jobs" -Completed
    Write-Progress -Id 9 -Activity "Starting Windows Update Services" -Completed
    Write-Progress -Id 10 -Activity "Forcing discovery" -Completed
}





# Funtion to uninstall Winget Packages
#---------------------------------------

<# .SYNOPSIS
    Uninstalls a package by its name using WinGet.

.DESCRIPTION
    The Uninstall-PackageByName function retrieves the list of installed packages using WinGet,
    searches for packages that match the provided name, and uninstalls them.

.PARAMETER packageName
    The name of the package to uninstall. This can be a partial name, and the function will
    match any installed packages that contain this string.

.EXAMPLE
    Uninstall-PackageByName -packageName "Visual Studio Code"
    This command will uninstall any installed packages that have "Visual Studio Code" in their name.

.NOTES
    This function requires WinGet to be installed and available in the system's PATH.

.OUTPUTS
    Outputs the status of the uninstallation process to the console.
    Function to uninstall a package by name#>
function Uninstall-PackageByName {
    param (
        [string]$packageName
    )

    # Special Package exceptions
    if ($packageName -eq "Power Automate") {
        Write-Output "Uninstalling Power Automate..."
        Start-Process -FilePath "winget" -ArgumentList "uninstall --id 9NFTCH6J7FHV -e" -NoNewWindow -Wait
    }

    # Get the list of installed packages as an object
    $installedPackages = Get-WinGetPackage

    # Find the correct package(s) matching the packageName
    $matchingPackages = $installedPackages | Where-Object { $_.Name -like "*$packageName*" }

    if ($matchingPackages) {
        foreach ($package in $matchingPackages) {
            $packageId = $package.Id
            Write-Output "Uninstalling package: $packageName with ID: $packageId"
            Start-Process -FilePath "winget" -ArgumentList "uninstall --id $packageId -e" -NoNewWindow -Wait
        }
    }
    else {
        Write-Output "Package $packageName not found."
    }
}

# Function to add the publisher to the trusted list 
#--------------------------------------------------
<# .SYNOPSIS
    Adds a specified publisher to the trusted list.

.DESCRIPTION
    The Add-PublisherToTrustedList function checks the current execution policy and sets it to RemoteSigned if it is not already Unrestricted or RemoteSigned.
    It then checks if the specified publisher is already in the trusted list and adds it if not. This ensures that scripts from the specified publisher can run without prompts.

.PARAMETER publisher
    The distinguished name (DN) of the publisher to add to the trusted list. This should include details such as CN, O, L, S, and C.

.EXAMPLE
    Add-PublisherToTrustedList -publisher "CN=Microsoft Corporation, O=Microsoft Corporation, L=Redmond, S=Washington, C=US"
    This command will add Microsoft Corporation to the trusted list, allowing scripts signed by this publisher to run without prompts.

.NOTES
    This function requires that the publisher's certificate is installed in the user's personal certificate store (Cert:\CurrentUser\My).
    The function will import the certificate to the TrustedPublisher store (Cert:\CurrentUser\TrustedPublisher).

.OUTPUTS
    Outputs the status of the operation to the console.
#>

function Add-PublisherToTrustedList {
    param (
        [string]$publisher
    )

    $executionPolicy = Get-ExecutionPolicy -Scope CurrentUser
    if ($executionPolicy -ne 'Unrestricted' -and $executionPolicy -ne 'RemoteSigned') {
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
    }

    $trustedPublisher = Get-ChildItem Cert:\CurrentUser\TrustedPublisher | Where-Object { $_.Subject -eq $publisher }
    if (-not $trustedPublisher) {
        Write-Host "Adding publisher to the trusted list..." -ForegroundColor Yellow
        $cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -eq $publisher }
        if ($cert) {
            Import-Certificate -FilePath $cert.PSPath -CertStoreLocation Cert:\CurrentUser\TrustedPublisher
            Write-Host "Publisher added to the trusted list." -ForegroundColor Green
        }
        else {
            Write-Warning "Publisher certificate not found. Please ensure the publisher certificate is installed."
        }
    }
    else {
        Write-Host "Publisher is already in the trusted list." -ForegroundColor Green
    }
}

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
    Write-Host "Removing Adobe Reader..."
        
    # Step 1: Uninstall Adobe Reader using the Adobe Reader removal tool
    try {
        $adobeRemovalToolPath = "$env:TEMP\AdobeAcroCleaner_DC2021.exe"
        if (-not (Test-Path $adobeRemovalToolPath)) {
            Write-Host "Downloading the Adobe Reader removal tool..." -ForegroundColor Yellow
            Invoke-WebRequest -Uri "https://www.adobe.com/support/downloads/product.jsp?product=1&platform=Windows" -OutFile $adobeRemovalToolPath
        }
        Write-Host "Running the Adobe Reader removal tool..." -ForegroundColor Yellow
        Start-Process -FilePath $adobeRemovalToolPath -ArgumentList "/silent" -NoNewWindow -Wait
        Write-Host "Adobe Reader removal tool has completed." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to run the Adobe Reader removal tool." -ForegroundColor Red
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
        Write-Host "Failed to uninstall Adobe Reader using winget." -ForegroundColor Red
    }
        
    # Step 3: Remove Adobe Reader-related folders
    $adobeReaderFolders = @(
        "$env:ProgramFiles\Adobe\Acrobat Reader DC",
        "$env:ProgramFiles (x86)\Adobe\Acrobat Reader DC",
        "$env:ProgramData\Adobe\Acrobat",
        "$env:LOCALAPPDATA\Adobe\Acrobat",
        "$env:APPDATA\Adobe\Acrobat"
    )
    foreach ($folder in $adobeReaderFolders) {
        try {
            Remove-Item -Recurse -Force -Path $folder
            Write-Host "Removed folder: $folder" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to remove folder: $folder" -ForegroundColor Red
        }
    }
        
    # Step 4: Remove Adobe Reader-related registry entries
    $adobeReaderRegistryPaths = @(
        "HKCU:\Software\Adobe\Acrobat Reader",
        "HKCU:\Software\Adobe\Acrobat Reader DC",
        "HKLM:\Software\Adobe\Acrobat Reader",
        "HKLM:\Software\Wow6432Node\Adobe\Acrobat Reader"
    )
    foreach ($regPath in $adobeReaderRegistryPaths) {
        try {
            Remove-Item -Recurse -Force -Path $regPath
            Write-Host "Removed registry path: $regPath" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to remove registry path: $regPath" -ForegroundColor Red
        }
    }

    Write-Host "Adobe Reader removal process is complete." -ForegroundColor Green
}  

###########################################
# Check Install Update and Import Modules #
###########################################

# Check if winget module is installed and up to date, if not install or update it
$moduleName = "Microsoft.WinGet.Client"
$module = Get-Module -Name $moduleName -ListAvailable -ErrorAction SilentlyContinue

if (-not $module) {
    Write-Host "Installing the $moduleName module..." -ForegroundColor Yellow
    try {
        Install-Module -Name $moduleName -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
        Write-Host "$moduleName module installed successfully." -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to install the $moduleName module"
    }
}
else {
    Write-Host "$moduleName module is already installed. Checking for updates..." -ForegroundColor Yellow
    try {
        $updateAvailable = Find-Module -Name $moduleName | Where-Object { $_.Version -gt $module.Version }
        if ($updateAvailable) {
            Write-Host "Updating the $moduleName module to the latest version..." -ForegroundColor Yellow
            Update-Module -Name $moduleName -Force -ErrorAction Stop
            Write-Host "$moduleName module updated successfully." -ForegroundColor Green
        }
        else {
            Write-Host "$moduleName module is up to date." -ForegroundColor Green
        }
    }
    catch {
        Write-Warning "Failed to check for updates or update the $moduleName module."
    }
}

# Add Microsoft Corporation to the trusted list
Add-PublisherToTrustedList -publisher "CN=Microsoft Corporation, O=Microsoft Corporation, L=Redmond, S=Washington, C=US"

# Import the winget module
try {
    Import-Module -Name Microsoft.WinGet.Client -ErrorAction Stop
    Write-Host "Microsoft.WinGet.Client module is imported." -ForegroundColor Green
}
catch {
    Write-Warning "Failed to import the Microsoft.WinGet.Client module. The Install Tab will not work properly."
}

##################################
# GUI Creation and Functionality #
##################################

# Import necessary .NET assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the form of the main GUI window
$form = New-Object System.Windows.Forms.Form
$form.Text = "ESS Tool Box (Beta) v0.1"
$form.Size = New-Object System.Drawing.Size(580, 470)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.MinimizeBox = $false

# Create a TabControl
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Size = New-Object System.Drawing.Size(550, 410)
$tabControl.Location = New-Object System.Drawing.Point(10, 10)
$form.Controls.Add($tabControl)

# Define Visual Elements for All Tabs
# Define column positions
$column1X = 20
$column2X = 200
$column3X = 380
#Define Section Length
$sectionLength = 520
#Define Install and Tweak Tab Buttons y position
$buttonY = 340

###################################################
# Install / Update Tab Creation and Functionality #
###################################################

# Create the Install tab
$tabInstall = New-Object System.Windows.Forms.TabPage
$tabInstall.Text = "Install / Update"
$tabControl.Controls.Add($tabInstall)

# Create checkboxes for packages in the Install tab
$checkboxAdobe = New-Object System.Windows.Forms.CheckBox
$checkboxAdobe.Text = "Adobe Reader"
$checkboxAdobe.Name = "Adobe Acrobat Reader 64-Bit"
$checkboxAdobe.AutoSize = $true
$checkboxAdobe.Location = New-Object System.Drawing.Point($column1X, 20)
$tabInstall.Controls.Add($checkboxAdobe)

$checkboxAdobeCloud = New-Object System.Windows.Forms.CheckBox
$checkboxAdobeCloud.Text = "Adobe Creative Cloud"
$checkboxAdobeCloud.Name = "Adobe Creative Cloud"
$checkboxAdobeCloud.AutoSize = $true
$checkboxAdobeCloud.Location = New-Object System.Drawing.Point($column1X, 50)
$tabInstall.Controls.Add($checkboxAdobeCloud)

$checkboxOffice = New-Object System.Windows.Forms.CheckBox
$checkboxOffice.Text = "Microsoft Office 365"
$checkboxOffice.Name = "Microsoft 365 Apps for Enterprise"
$checkboxOffice.AutoSize = $true
$checkboxOffice.Location = New-Object System.Drawing.Point($column1X, 80)
$tabInstall.Controls.Add($checkboxOffice)

$checkboxOneNote = New-Object System.Windows.Forms.CheckBox
$checkboxOneNote.Text = "Microsoft OneNote UWP (msstore)"
$checkboxOneNote.Name = "Microsfot OneNote"
$checkboxOneNote.AutoSize = $true
$checkboxOneNote.Location = New-Object System.Drawing.Point($column1X, 110)
$tabInstall.Controls.Add($checkboxOneNote)

$checkboxTeams = New-Object System.Windows.Forms.CheckBox
$checkboxTeams.Text = "Microsoft Teams"
$checkboxTeams.Name = "Microsoft Teams"
$checkboxTeams.AutoSize = $true
$checkboxTeams.Location = New-Object System.Drawing.Point($column1X, 140)
$tabInstall.Controls.Add($checkboxTeams)

$checkboxNetFrameworks = New-Object System.Windows.Forms.CheckBox
$checkboxNetFrameworks.Text = ".NET All Versions"
$checkboxNetFrameworks.Name = "Microsoft .Net Runtime"
$checkboxNetFrameworks.AutoSize = $true
$checkboxNetFrameworks.Location = New-Object System.Drawing.Point($column1X, 170)
$tabInstall.Controls.Add($checkboxNetFrameworks)

$checkboxPowerAutomate = New-Object System.Windows.Forms.CheckBox
$checkboxPowerAutomate.Text = "Power Automate"
$checkboxPowerAutomate.Name = "Power Automate"
$checkboxPowerAutomate.AutoSize = $true
$checkboxPowerAutomate.Location = New-Object System.Drawing.Point($column1X, 200)
$tabInstall.Controls.Add($checkboxPowerAutomate)

$checkboxPowerToys = New-Object System.Windows.Forms.CheckBox
$checkboxPowerToys.Text = "PowerToys"
$checkboxPowerToys.Name = "PowerToys"
$checkboxPowerToys.AutoSize = $true
$checkboxPowerToys.Location = New-Object System.Drawing.Point($column1X, 230)
$tabInstall.Controls.Add($checkboxPowerToys)

$checkboxQuickAssist = New-Object System.Windows.Forms.CheckBox
$checkboxQuickAssist.Text = "Quick Assist"
$checkboxQuickAssist.Name = "Quick Assist"
$checkboxQuickAssist.AutoSize = $true
$checkboxQuickAssist.Location = New-Object System.Drawing.Point($column1X, 260)
$tabInstall.Controls.Add($checkboxQuickAssist)

$checkboxSurfaceDiagnosticToolkit = New-Object System.Windows.Forms.CheckBox
$checkboxSurfaceDiagnosticToolkit.Text = "Surface Diagnostic Toolkit"
$checkboxSurfaceDiagnosticToolkit.Name = "Surface Diagnostic Toolkit"
$checkboxSurfaceDiagnosticToolkit.AutoSize = $true
$checkboxSurfaceDiagnosticToolkit.Location = New-Object System.Drawing.Point($column2X, 20)
$tabInstall.Controls.Add($checkboxSurfaceDiagnosticToolkit)

$checkboxVisio = New-Object System.Windows.Forms.CheckBox
$checkboxVisio.Text = "Visio Viewer 2016"
$checkboxVisio.Name = "Microsoft VisioViewer"
$checkboxVisio.AutoSize = $true
$checkboxVisio.Location = New-Object System.Drawing.Point($column2X, 80)
$tabInstall.Controls.Add($checkboxVisio)

$checkboxRemoteDesktop = New-Object System.Windows.Forms.CheckBox
$checkboxRemoteDesktop.Text = "Remote Desktop"
$checkboxRemoteDesktop.Name = "Microsoft Remote Desktop"
$checkboxRemoteDesktop.AutoSize = $true
$checkboxRemoteDesktop.Location = New-Object System.Drawing.Point($column2X, 50)
$tabInstall.Controls.Add($checkboxRemoteDesktop)

# Create a checkbox to show in the command promt all pacakges installed
$checkboxShowInstalled = New-Object System.Windows.Forms.CheckBox
$checkboxShowInstalled.Text = "Show Installed Packages"
$checkboxShowInstalled.Name = "ShowInstalled"
$checkboxShowInstalled.AutoSize = $true
$checkboxShowInstalled.Location = New-Object System.Drawing.Point($column3X, $buttonY)
$tabInstall.Controls.Add($checkboxShowInstalled)

# Create an Install button in the Install tab
$buttonInstall = New-Object System.Windows.Forms.Button
$buttonInstall.Text = "Install"
$buttonInstall.AutoSize = $true
$buttonInstall.Location = New-Object System.Drawing.Point(20, $buttonY)
$tabInstall.Controls.Add($buttonInstall)

# Define the action for the Install button
$buttonInstall.Add_Click({
        if ($checkboxOffice.Checked) {
            Start-Process "winget" -ArgumentList "install --id Microsoft.Office -e --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait
        }
        if ($checkboxPowerToys.Checked) {
            Start-Process "winget" -ArgumentList "install --id Microsoft.PowerToys -e --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait
        }
        if ($checkboxTeams.Checked) {
            Start-Process "winget" -ArgumentList "install --id Microsoft.Teams -e --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait
        }
        if ($checkboxOneNote.Checked) {
            Start-Process "winget" -ArgumentList "install --id XPFFZHVGQWWLHB -e --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait
        }
        if ($checkboxAdobe.Checked) {
            Start-Process "winget" -ArgumentList "install --id Adobe.Acrobat.Reader.64-Bit -e --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait
        }
        if ($checkboxAdobeCloud.Checked) {
            Start-Process "winget" -ArgumentList "install --id Adobe.CreativeCloud -e --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait
        }
        if ($checkboxPowerAutomate.Checked) {
            Start-Process "winget" -ArgumentList "install --id 9NFTCH6J7FHV -e --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait
        }
        if ($checkboxVisio.Checked) {
            Start-Process "winget" -ArgumentList "install --id Microsoft.VisioViewer -e --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait
        }
        if ($checkboxNetFrameworks.Checked) {
            Start-Process "winget" -ArgumentList "install --id Microsoft.DotNet.DesktopRuntime.3_1 -e --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait
            Start-Process "winget" -ArgumentList "install --id Microsoft.DotNet.DesktopRuntime.5 -e --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait
            Start-Process "winget" -ArgumentList "install --id Microsoft.DotNet.DesktopRuntime.6 -e --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait
            Start-Process "winget" -ArgumentList "install --id Microsoft.DotNet.DesktopRuntime.7 -e --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait
            Start-Process "winget" -ArgumentList "install --id Microsoft.DotNet.DesktopRuntime.8 -e --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait
        }
        if ($checkboxQuickAssist.Checked) {
            Start-Process "winget" -ArgumentList "install --id 9P7BP5VNWKX5 -e --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait
        }
        if ($checkboxSurfaceDiagnosticToolkit.Checked) {
            Start-Process "winget" -ArgumentList "install --id 9NF1MR6C60ZF -e --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait
        }
        if($checkboxRemoteDesktop.Checked){
            Start-Process "winget" -ArgumentList "install --id 9WZDNCRFJ3PS -e --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait
        }
        [System.Windows.Forms.MessageBox]::Show("Selected packages have been installed.")
    })

# Create an Uninstall button in the Install tab
$buttonUninstall = New-Object System.Windows.Forms.Button
$buttonUninstall.Text = "Uninstall"
$buttonUninstall.AutoSize = $true
$buttonUninstall.Location = New-Object System.Drawing.Point(120, $buttonY)
$tabInstall.Controls.Add($buttonUninstall)

# Define the action for the Uninstall button
$buttonUninstall.Add_Click({

        # Loop through each package and uninstall it
        $packagesToUninstall = @()
        foreach ($control in $tabInstall.Controls) {
            if ($control -is [System.Windows.Forms.CheckBox] -and $control.Checked) {
                $packagesToUninstall += $control.Name
            }
        }
        foreach ($package in $packagesToUninstall) {
            Uninstall-PackageByName -packageName $package
        }

        [System.Windows.Forms.MessageBox]::Show("Selected packages have been uninstalled.")
    })

# Create a Get Installed Packages button in the Install tab
$buttonGetPackages = New-Object System.Windows.Forms.Button
$buttonGetPackages.Text = "Get Installed"
$buttonGetPackages.AutoSize = $true
$buttonGetPackages.Location = New-Object System.Drawing.Point(220, $buttonY)
$tabInstall.Controls.Add($buttonGetPackages)

# Define the action for the Get Packages button
$buttonGetPackages.Add_Click({
        # Check if the showInstalled checkbox is checked

        $showInstalledCheckbox = $tabInstall.Controls | Where-Object { $_.Name -eq "showInstalled" -and $_.Checked }
        if ($showInstalledCheckbox) {
            Start-Process "winget" -ArgumentList "list" -NoNewWindow -Wait
        }

        # Run the winget list command and capture the output directly
        $output = & winget list
        # Split the output into lines and skip the first two lines (headers)
        $outputLines = $output -split '\r?\n'

        # Iterate through each control in the install tab
        foreach ($control in $tabInstall.Controls) {
            if ($control -is [System.Windows.Forms.CheckBox]) {
                $checkboxName = $control.Name
                # Check if the checkbox name is present in the output lines
                if ($outputLines -match $checkboxName) {
                    $control.Checked = $true
                }
                else {
                    $control.Checked = $false
                }
            }
        }
    })

#######################################
# Tweak Tab Creation and Functionality #
#######################################

# Create the Tweak tab
$tabTweak = New-Object System.Windows.Forms.TabPage
$tabTweak.Text = "Tweak"
$tabControl.Controls.Add($tabTweak)

# Create controls for the Tweak tab
$checkboxRightClickEndTask = New-Object System.Windows.Forms.CheckBox
$checkboxRightClickEndTask.Text = "Enable End Task With Right Click"
$checkboxRightClickEndTask.Name = "EnableRightClickEndTask"
$checkboxRightClickEndTask.AutoSize = $true
$checkboxRightClickEndTask.Location = New-Object System.Drawing.Point($column1X, 110)
$tabTweak.Controls.Add($checkboxRightClickEndTask)

$checkboxRunDiskCleanup = New-Object System.Windows.Forms.CheckBox
$checkboxRunDiskCleanup.Text = "Run Disk Cleanup"
$checkboxRunDiskCleanup.Name = "RunDiskCleanup"
$checkboxRunDiskCleanup.AutoSize = $true
$checkboxRunDiskCleanup.Location = New-Object System.Drawing.Point($column1X, 230)
$tabTweak.Controls.Add($checkboxRunDiskCleanup)

$checkboxDetailedBSOD = New-Object System.Windows.Forms.CheckBox
$checkboxDetailedBSOD.Text = "Enable Detailed BSOD Information"
$checkboxDetailedBSOD.Name = "EnableDetailedBSOD"
$checkboxDetailedBSOD.AutoSize = $true
$checkboxDetailedBSOD.Location = New-Object System.Drawing.Point($column1X, 80)
$tabTweak.Controls.Add($checkboxDetailedBSOD)

$checkboxVerboseLogon = New-Object System.Windows.Forms.CheckBox
$checkboxVerboseLogon.Text = "Enable Verbose Logon Messages"
$checkboxVerboseLogon.Name = "EnableVerboseLogon"
$checkboxVerboseLogon.AutoSize = $true
$checkboxVerboseLogon.Location = New-Object System.Drawing.Point($column1X, 170)
$tabTweak.Controls.Add($checkboxVerboseLogon)

$checkboxDeleteTempFiles = New-Object System.Windows.Forms.CheckBox
$checkboxDeleteTempFiles.Text = "Delete Temporary Files"
$checkboxDeleteTempFiles.Name = "DeleteTempFiles"
$checkboxDeleteTempFiles.AutoSize = $true
$checkboxDeleteTempFiles.Location = New-Object System.Drawing.Point($column1X, 20)
$tabTweak.Controls.Add($checkboxDeleteTempFiles)

$checkboxClassicRightClickMenu = New-Object System.Windows.Forms.CheckBox
$checkboxClassicRightClickMenu.Text = "Enable Classic Right Click Menu"
$checkboxClassicRightClickMenu.Name = "EnableClassicRightClickMenu"
$checkboxClassicRightClickMenu.AutoSize = $true
$checkboxClassicRightClickMenu.Location = New-Object System.Drawing.Point($column1X, 50)
$tabTweak.Controls.Add($checkboxClassicRightClickMenu)

$checkboxGodMode = New-Object System.Windows.Forms.CheckBox
$checkboxGodMode.Text = "Enable God Mode"
$checkboxGodMode.Name = "EnableGodMode"
$checkboxGodMode.AutoSize = $true
$checkboxGodMode.Location = New-Object System.Drawing.Point($column1X, 140)
$tabTweak.Controls.Add($checkboxGodMode)

$checkboxOptimizeDrives = New-Object System.Windows.Forms.CheckBox
$checkboxOptimizeDrives.Text = "Optimize Drives"
$checkboxOptimizeDrives.Name = "OptimizeDrives"
$checkboxOptimizeDrives.AutoSize = $true
$checkboxOptimizeDrives.Location = New-Object System.Drawing.Point($column1X, 200)
$tabTweak.Controls.Add($checkboxOptimizeDrives)

# Create Apply and Undo buttons in the Tweak tab
$buttonApply = New-Object System.Windows.Forms.Button
$buttonApply.Text = "Apply"
$buttonApply.AutoSize = $true
$buttonApply.Location = New-Object System.Drawing.Point(20, $buttonY)
$tabTweak.Controls.Add($buttonApply)

$buttonUndo = New-Object System.Windows.Forms.Button
$buttonUndo.Text = "Undo"
$buttonUndo.AutoSize = $true
$buttonUndo.Location = New-Object System.Drawing.Point(120, $buttonY)
$tabTweak.Controls.Add($buttonUndo)

#Create a System Performance Button
$buttonSystemPerformance = New-Object System.Windows.Forms.Button  
$buttonSystemPerformance.Text = "System Performance"
$buttonSystemPerformance.AutoSize = $true
$buttonSystemPerformance.Location = New-Object System.Drawing.Point(220, $buttonY)
$tabTweak.Controls.Add($buttonSystemPerformance)

# Define the action for the Apply button
$buttonApply.Add_Click({
        if ($checkboxRightClickEndTask.Checked) {
            # Add registry key to enable right click end task
            $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarDeveloperSettings"
            $regName = "TaskbarEndTask"
            $regValue = 1
            #Ensure the registry path exists
            if (-not (Test-Path $regPath)) {
                New-Item -Path $regPath -Force | Out-Null
            }
            #Set the registry value, creating it if it doesn't exist
            New-ItemProperty -Path $regPath -Name $regName -Value $regValue -Force | Out-Null
        }
        if ($checkboxRunDiskCleanup.Checked) {
            # Run disk cleanup
            Start-Process "cleanmgr.exe" -ArgumentList "/sagerun:1" -NoNewWindow -Wait
        
            # Run DISM command with /StartComponentCleanup
            Start-Process "dism.exe" -ArgumentList "/Online /Cleanup-Image /StartComponentCleanup" -NoNewWindow -Wait
        }
        if ($checkboxDetailedBSOD.Checked) {
            # Enable detailed BSOD information
            Write-Host "Enabling detailed BSOD information..." -ForegroundColor Green
        
            try {
                # Define the registry path
                $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Control\CrashControl"
            
                # Check if the registry path exists, and create it if it does not
                if (-not (Test-Path $registryPath)) {
                    New-Item -Path $registryPath -Force
                    Write-Output "Created registry path: $registryPath"
                }
            
                # Set the registry key to enable detailed BSOD information
                Set-ItemProperty -Path $registryPath -Name "DisplayParameters" -Value 1 -Force
                Write-Host "Detailed BSOD information has been enabled." -ForegroundColor Green
            }
            catch {
                Write-Host "Failed to enable detailed BSOD information: $_" -ForegroundColor Red
            }
        }
        if ($checkboxVerboseLogon.Checked) {
            # Enable verbose logon messages
            Write-Output "Enabling verbose logon messages..."
            # Add your code here
            try {
                # Define the registry path
                $registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
            
                # Check if the registry path exists, and create it if it does not
                if (-not (Test-Path $registryPath)) {
                    New-Item -Path $registryPath -Force
                    Write-Output "Created registry path: $registryPath"
                }
            
                # Set the registry key to enable verbose logon messages
                Set-ItemProperty -Path $registryPath -Name "VerboseStatus" -Value 1 -Force
                Write-Host "Verbose logon messages have been enabled." -ForegroundColor Green
            }
            catch {
                Write-Host "Failed to enable verbose logon messages: $_" -ForegroundColor Red
            }
        }
        if ($checkboxDeleteTempFiles.Checked) {
            # Delete temporary files
            Write-Host "Deleting temporary files..." -ForegroundColor Green

            # Check if the C:\Windows\Temp path exists
            if (Test-Path "C:\Windows\Temp") {
                Get-ChildItem -Path "C:\Windows\Temp" *.* -Recurse | Remove-Item -Force -Recurse
            }
            else {
                Write-Host "Path C:\Windows\Temp does not exist." -ForegroundColor Red
            }

            # Check if the $env:TEMP path exists
            if (Test-Path $env:TEMP) {
                Get-ChildItem -Path $env:TEMP *.* -Recurse | Remove-Item -Force -Recurse
            }
            else {
                Write-Host "Path $env:TEMP does not exist." -ForegroundColor Red
            }

        }
        if ($checkboxClassicRightClickMenu.Checked) {
            # Enable Classic Right Click Menu
            New-Item -Path "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}" -Name "InprocServer32" -Force -Value ""
            Write-Host "Classic Right Click Menu has been enabled." -ForegroundColor Green

            # Restart explorer.exe
            Write-Host "Restarting explorer.exe ..." -ForegroundColor Green
            $process = Get-Process -Name "explorer"
            Stop-Process -InputObject $process
        }
        if ($checkboxGodMode.Checked) {
            $desktopPath = [System.Environment]::GetFolderPath('Desktop')
            $godModePath = "$desktopPath\GodMode.{ED7BA470-8E54-465E-825C-99712043E01C}"
            if (-not (Test-Path $godModePath)) {
                try {
                    New-Item -Path $godModePath -ItemType Directory -Force | Out-Null
                    Write-Host "God Mode has been enabled on the desktop." -ForegroundColor Green
                }
                catch {
                    Write-Host "Failed to enable God Mode. Error: $_" -ForegroundColor Red
                }
            }
            else {
                Write-Host "God Mode is already enabled on the desktop." -ForegroundColor Yellow
            }
        }
        if ($checkboxOptimizeDrives.Checked) {
            # Optimize Drives
            Write-Host "Optimizing drives..." -ForegroundColor Green
            Get-Volume | Where-Object { $_.DriveLetter } | ForEach-Object { 
                try {
                    Optimize-Volume -DriveLetter $_.DriveLetter -ReTrim -Verbose
                    Write-Host "Drive $($_.DriveLetter) has been optimized." -ForegroundColor Green
                }
                catch {
                    Write-Host "Failed to optimize drive $($_.DriveLetter). Error: $_" -ForegroundColor Yellow
                }
            }
            Write-Host "Drive optimization process completed." -ForegroundColor Green
        }
        [System.Windows.Forms.MessageBox]::Show("Selected tweaks have been applied.")
    })

# Define the action for the Undo button
$buttonUndo.Add_Click({
        if ($checkboxRightClickEndTask.Checked) {
            # Remove registry key to disable right click end task
            $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarDeveloperSettings"
            $regName = "TaskbarEndTask"
            $regValue = 0
        
            #Ensure the registry path exists
            if (-not (Test-Path $regPath)) {
                New-Item -Path $regPath -Force | Out-Null
            }
            #Remove the registry value, creating it if it doesn't exist
            New-ItemProperty -Path $regPath -Name $regName -Value $regValue -Force | Out-Null
        }
        if ($checkboxRunDiskCleanup.Checked) {
            # Undo disk cleanup
            Write-Output "Nothing to do here..."
        }
        if ($checkboxDetailedBSOD.Checked) {
            # Disable detailed BSOD information
            Write-Host "Disabling detailed BSOD information..." -ForegroundColor Green
    
            try {
                # Define the registry path
                $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Control\CrashControl"
        
                # Check if the registry path exists
                if (Test-Path $registryPath) {
                    # Set the registry key to disable detailed BSOD information
                    Set-ItemProperty -Path $registryPath -Name "DisplayParameters" -Value 0 -Force
                    Write-Host "Detailed BSOD information has been disabled." -ForegroundColor Green
                }
                else {
                    Write-Host "Registry path does not exist: $registryPath" -ForegroundColor Red
                }
            }
            catch {
                Write-Host "Failed to disable detailed BSOD information: $_" -ForegroundColor Red
            }
        }
        if ($checkboxVerboseLogon.Checked) {
            # Disable verbose logon messages
            Write-Output "Disabling verbose logon messages..."
            # Add your code here
            try {
                # Define the registry path
                $registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
            
                # Check if the registry path exists
                if (Test-Path $registryPath) {
                    # Remove the registry key to disable verbose logon messages
                    Remove-ItemProperty -Path $registryPath -Name "VerboseStatus" -Force
                    Write-Host "Verbose logon messages have been disabled." -ForegroundColor Green
                }
                else {
                    Write-Host "Registry path does not exist: $registryPath" -ForegroundColor Red
                }
            }
            catch {
                Write-Host "Failed to disable verbose logon messages: $_" -ForegroundColor Red
            }
        }
        if ($checkboxDeleteTempFiles.Checked) {
            # Undo delete temporary files
            Write-Output "Nothing to do here..."
            # Add your code here
        }
        if ($checkboxClassicRightClickMenu.Checked) {
            # Disable Classic Right Click Menu
            Remove-Item -Path "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}" -Recurse -Confirm:$false -Force
            Write-Host "Classic Right Click Menu has been disabled." -ForegroundColor Green

            # Restart explorer.exe
            Write-Host "Restarting explorer.exe ..." -ForegroundColor Green
            $process = Get-Process -Name "explorer"
            Stop-Process -InputObject $process
        }
        if ($checkboxGodMode.Checked) {
            # Disable God Mode
            $desktopPath = [System.Environment]::GetFolderPath('Desktop')
            $godModePath = "$desktopPath\GodMode.{ED7BA470-8E54-465E-825C-99712043E01C}"
            if (Test-Path $godModePath) {
                try {
                    Remove-Item -Path $godModePath -Recurse -Force
                    Write-Host "God Mode has been disabled." -ForegroundColor Green
                }
                catch {
                    Write-Host "Failed to disable God Mode. Error: $_" -ForegroundColor Red
                }
            }
            else {
                Write-Host "God Mode folder does not exist." -ForegroundColor Yellow
            }
        }
        if ($checkboxOptimizeDrives.Checked) {
            # Undo Optimize Drives
            Write-Output "Nothing to do here..."
        }

        [System.Windows.Forms.MessageBox]::Show("Selected tweaks have been undone.")
    })

# Define the action for the System Performance button
$buttonSystemPerformance.Add_Click({
        # Open System Performance settings
        Write-Host "Opening System Performance settings..." -ForegroundColor Green
        Start-Process "SystemPropertiesPerformance.exe"
    })

######################################
# Fix Tab Creation and Functionality #
######################################

# Create the Fix tab
$tabFix = New-Object System.Windows.Forms.TabPage
$tabFix.Text = "Fix"
$tabControl.Controls.Add($tabFix)

###############
# Apps section#
###############

# Create controls for the Fix tab
$sectionApps = New-Object System.Windows.Forms.GroupBox
$sectionApps.Text = "Apps"
$sectionApps.Size = New-Object System.Drawing.Size($sectionLength, 100)
$sectionApps.Location = New-Object System.Drawing.Point($column1X, 20)
$tabFix.Controls.Add($sectionApps)

# Create a hyperlink to remove Adobe Cloud
$linkRemoveAdobeCloud = New-Object System.Windows.Forms.LinkLabel
$linkRemoveAdobeCloud.Text = "Remove Adobe Creative Cloud"
$linkRemoveAdobeCloud.AutoSize = $true
$linkRemoveAdobeCloud.Location = New-Object System.Drawing.Point($column1X, 30)
$linkRemoveAdobeCloud.Add_LinkClicked({
        # Remove Adobe Cloud
        Write-Host "Removing Adobe Cloud..."
        # Code snipet from https://github.com/ChrisTitusTech/winutil/blob/main/docs/dev/features/Fixes/RunAdobeCCCleanerTool.md
        [string]$url = "https://swupmf.adobe.com/webfeed/CleanerTool/win/AdobeCreativeCloudCleanerTool.exe"

        Write-Host "The Adobe Creative Cloud Cleaner tool is hosted at"
        Write-Host "$url"

        try {
            # Don't show the progress because it will slow down the download speed
            $ProgressPreference = 'SilentlyContinue'

            Invoke-WebRequest -Uri $url -OutFile "$env:TEMP\AdobeCreativeCloudCleanerTool.exe" -UseBasicParsing -ErrorAction SilentlyContinue -Verbose

            # Revert back the ProgressPreference variable to the default value since we got the file desired
            $ProgressPreference = 'Continue'

            Start-Process -FilePath "$env:TEMP\AdobeCreativeCloudCleanerTool.exe" -Wait -ErrorAction SilentlyContinue -Verbose
        }
        catch {
            Write-Error $_.Exception.Message
        }
        finally {
            if (Test-Path -Path "$env:TEMP\AdobeCreativeCloudCleanerTool.exe") {
                Write-Host "Cleaning up..."
                Remove-Item -Path "$env:TEMP\AdobeCreativeCloudCleanerTool.exe" -Verbose
            }
        }
    })
$sectionApps.Controls.Add($linkRemoveAdobeCloud)

# Create a hyperlink to remove Adobe Reader
$linkRemoveAdobeReader = New-Object System.Windows.Forms.LinkLabel
$linkRemoveAdobeReader.Text = "Remove Adobe Reader"
$linkRemoveAdobeReader.AutoSize = $true
$linkRemoveAdobeReader.Location = New-Object System.Drawing.Point($column1X, 60)
$linkRemoveAdobeReader.Add_LinkClicked({
        # Remove Adobe Reader
        Write-Output "Removing Adobe Reader..."
        # Call the function to remove Adobe Reader
        Remove-AdobeReader
    })
$sectionApps.Controls.Add($linkRemoveAdobeReader)

# Create a hyperlink to reset Edge Browser Cache
$linkResetEdgeCache = New-Object System.Windows.Forms.LinkLabel
$linkResetEdgeCache.Text = "Reset Edge Browser Cache"
$linkResetEdgeCache.AutoSize = $true
$linkResetEdgeCache.Location = New-Object System.Drawing.Point($column2X, 30)
$linkResetEdgeCache.Add_LinkClicked({
    try {
        # Open Edge browser
        $edgeProcess = Start-Process "msedge" -ArgumentList "about:blank" -PassThru

        # Wait for Edge to open
        Start-Sleep -Seconds 3

        # Simulate key presses to navigate to the settings page for clearing browsing data
        $shell = New-Object -ComObject "WScript.Shell"
        $shell.AppActivate($edgeProcess.Id)
        Start-Sleep -Milliseconds 500
        $shell.SendKeys("^+{DEL}")  # Ctrl+Shift+Delete to open the Clear browsing data dialog
        Start-Sleep -Milliseconds 500
        $shell.SendKeys("{ENTER}")  # Press Enter to confirm

        Write-Host "Edge settings page for clearing browsing data has been opened." -ForegroundColor Green
    } catch {
        Write-Host "Failed to open Edge settings page. Error: $_" -ForegroundColor Red
    }
})
$sectionApps.Controls.Add($linkResetEdgeCache)

# Create a hyperlink to fully reset Edge Browser
$linkResetEdge = New-Object System.Windows.Forms.LinkLabel
$linkResetEdge.Text = "Reset Edge Browser"
$linkResetEdge.AutoSize = $true
$linkResetEdge.Location = New-Object System.Drawing.Point($column2X, 60)
$linkResetEdge.Add_LinkClicked({
    try {
        # Show a message box to advise the user
        $result = [System.Windows.Forms.MessageBox]::Show(
            "This action will fully remove Microsoft Edge and all cached files from the system. Do you want to proceed?",
            "Warning",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            # Step 1: Reset Edge profile
            Write-Host "Resetting Edge browser profile..."
            Start-Process "msedge" -ArgumentList "--reset-profile" -Wait
            Write-Host "Edge browser profile has been reset." -ForegroundColor Green

            # Step 2: Remove Edge cache and temporary files
            Write-Host "Removing Edge cache and temporary files..."
            $edgeCachePaths = @(
                "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache",
                "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Code Cache",
                "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\GPUCache",
                "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Media Cache",
                "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\ShaderCache"
            )
            foreach ($path in $edgeCachePaths) {
                if (Test-Path $path) {
                    Remove-Item -Path $path -Recurse -Force
                    Write-Host "Removed: $path" -ForegroundColor Green
                } else {
                    Write-Host "Path not found: $path" -ForegroundColor Yellow
                }
            }

            # Step 3: Uninstall Edge
            Write-Host "Uninstalling Edge browser..."
            $edgeUninstallPath = "$env:ProgramFiles (x86)\Microsoft\Edge\Application\msedge.exe"
            if (Test-Path $edgeUninstallPath) {
                Start-Process $edgeUninstallPath -ArgumentList "--uninstall --system-level --force-uninstall" -Wait
                Write-Host "Edge browser has been uninstalled." -ForegroundColor Green
            } else {
                Write-Host "Edge uninstall path not found." -ForegroundColor Red
            }

            # Step 4: Reinstall Edge using winget
            Write-Host "Reinstalling Edge browser using winget..."
            Start-Process "winget" -ArgumentList "install --id Microsoft.Edge -e --source winget" -Wait
            Write-Host "Edge browser has been reinstalled." -ForegroundColor Green
        } else {
            Write-Host "Action canceled by the user." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Failed to reset Edge browser. Error: $_" -ForegroundColor Red
    }
})
$sectionApps.Controls.Add($linkResetEdge)

#################
# System section#
#################

$sectionSystem = New-Object System.Windows.Forms.GroupBox
$sectionSystem.Text = "System"
$sectionSystem.Size = New-Object System.Drawing.Size($sectionLength, 100)
$sectionSystem.Location = New-Object System.Drawing.Point($column1X, 260)
$tabFix.Controls.Add($sectionSystem)

# Create a hyperlink to reset Windows Update
$linkResetWinUpdate = New-Object System.Windows.Forms.LinkLabel
$linkResetWinUpdate.Text = "Reset Windows Update"
$linkResetWinUpdate.AutoSize = $true
$linkResetWinUpdate.Location = New-Object System.Drawing.Point($column1X, 30)
$linkResetWinUpdate.Add_LinkClicked({
        # Reset Windows Update
        Write-Output "Resetting Windows Update..."
        # Add your code here
        $fixWindowsUpdate = [System.Windows.Forms.MessageBox]::Show("We will attempt to fix Windows Update service. Do you want to run the fix in aggressive mode?", "Fix Windows Update", "YesNoCancel", "Question")

        if ($fixWindowsUpdate -eq "Yes") {
            Invoke-FixesWUpdate -Aggressive $true
        }
        elseif ($fixWindowsUpdate -eq "No") {
            Invoke-FixesWUpdate -Aggressive $false
        }
        else {
            # User clicked Cancel, do nothing
        }
    })
$sectionSystem.Controls.Add($linkResetWinUpdate)

# Create a hyperlink to reset network
$linkResetNetwork = New-Object System.Windows.Forms.LinkLabel
$linkResetNetwork.Text = "Reset Network"
$linkResetNetwork.AutoSize = $true
$linkResetNetwork.Location = New-Object System.Drawing.Point($column1X, 60)
$linkResetNetwork.Add_LinkClicked({
        # Reset network
        Write-Host "Resetting network..."

        # Reset network using netsh
        netsh winsock reset | Out-Null
        netsh int ip reset | Out-Null
        netsh interface ipv4 reset | Out-Null
        netsh interface ipv6 reset | Out-Null
        netsh interface reset all | Out-Null

        # Reset Wi-Fi card
        $wifiAdapter = Get-NetAdapter | Where-Object { $_.InterfaceDescription -like "*Wi-Fi*" }
        $wifiAdapter | Disable-NetAdapter | Out-Null
        $wifiAdapter | Enable-NetAdapter | Out-Null
        $wifiAdapter | Enable-NetAdapter | Out-Null

        # Display message to reboot device
        [System.Windows.Forms.MessageBox]::Show("Please reboot your device to complete the network reset process.", "Network Reset", "OK", "Information")

        # Network reset completed
        Write-Host "Network Reset Completed!" -ForegroundColor Green
        # Add your code here

    })
$sectionSystem.Controls.Add($linkResetNetwork)

# Create a hyperlink to run system repair
$linkSysRepair = New-Object System.Windows.Forms.LinkLabel
$linkSysRepair.Text = "System Repair"
$linkSysRepair.AutoSize = $true
$linkSysRepair.Location = New-Object System.Drawing.Point($column2X, 30)
$linkSysRepair.Add_LinkClicked({
        # Run system repair
        Write-Output "Running system repair..."
        <#  From ChrisTitusTech/winutil repo github.com/ChrisTitusTech/winutil/blob/main/docs/dev/features/Fixes/RunSystemRepair.md
    System Repair Steps:
        1. Chkdsk    - Fixes disk and filesystem corruption
        2. SFC Run 1 - Fixes system file corruption, and fixes DISM if it was corrupted
        3. DISM      - Fixes system image corruption, and fixes SFC's system image if it was corrupted
        4. SFC Run 2 - Fixes system file corruption, this time with an almost guaranteed uncorrupted system 
    #>
        Start-Process PowerShell -ArgumentList "Write-Host '(1/4) Chkdsk' -ForegroundColor Green; Chkdsk /scan;
    Write-Host '`n(2/4) SFC - 1st scan' -ForegroundColor Green; sfc /scannow;
    Write-Host '`n(3/4) DISM' -ForegroundColor Green; DISM /Online /Cleanup-Image /Restorehealth;
    Write-Host '`n(4/4) SFC - 2nd scan' -ForegroundColor Green; sfc /scannow;
    Read-Host '`nPress Enter to Continue'" -verb runas
    })
$sectionSystem.Controls.Add($linkSysRepair)

######################
# Office Apps section#
######################

#Create a group box for Office Apps
$sectionOfficeApps = New-Object System.Windows.Forms.GroupBox
$sectionOfficeApps.Text = "Office Apps"
$sectionOfficeApps.Size = New-Object System.Drawing.Size($sectionLength, 100)
$sectionOfficeApps.Location = New-Object System.Drawing.Point($column1X, 140)
$tabFix.Controls.Add($sectionOfficeApps)

# Create a hyperlink to rebuild .OST file
$linkRebuildOST = New-Object System.Windows.Forms.LinkLabel
$linkRebuildOST.Text = "Rebuild .OST file"
$linkRebuildOST.AutoSize = $true
$linkRebuildOST.Location = New-Object System.Drawing.Point($column1X, 30)
$linkRebuildOST.Add_LinkClicked({
        # Prompt the user for confirmation
        $result = [System.Windows.Forms.MessageBox]::Show("This action will force quit Outlook and Teams. Please save any important work before proceeding. Do you want to continue?", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            # Check if Outlook and Teams are running
            $outlookRunning = Get-Process -Name Outlook -ErrorAction SilentlyContinue
            $teamsRunning = Get-Process -Name Teams -ErrorAction SilentlyContinue

            # Terminate Outlook and Teams
            if ($outlookRunning) {
                Stop-Process -Name Outlook -Force
            }
            if ($teamsRunning) {
                Stop-Process -Name Teams -Force
            }

            # Path to the Outlook .ost files
            $ostPath = "$env:LOCALAPPDATA\Microsoft\Outlook"

            # Check if the path exists
            if (Test-Path $ostPath) {
                # Remove all .ost files in the current user's AppData folder
                $ostFiles = Get-ChildItem -Path $ostPath -Filter *.ost -Recurse -ErrorAction SilentlyContinue
                if ($ostFiles) {
                    foreach ($file in $ostFiles) {
                        Remove-Item -Path $file.FullName -Force
                    }
                    Write-Host "All .ost files have been removed." -ForegroundColor Green
                }
                else {
                    Write-Host "No .ost files found to rebuild." -ForegroundColor Red
                }
            }
            else {
                Write-Host "The path to .ost files does not exist." -ForegroundColor Red
            }

            # Reopen Outlook and Teams if they were running before
            if ($outlookRunning) {
                Start-Process -Name Outlook
            }
            if ($teamsRunning) {
                Start-Process -Name Teams
            }

            [System.Windows.Forms.MessageBox]::Show("The process is complete. Outlook and Teams have been restarted if they were running before.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })
$sectionOfficeApps.Controls.Add($linkRebuildOST)

# Create a hyperlink for Missing Teams Add-In
$linkMissingTeamsAddIn = New-Object System.Windows.Forms.LinkLabel
$linkMissingTeamsAddIn.Text = "Missing Teams Add-In"
$linkMissingTeamsAddIn.AutoSize = $true
$linkMissingTeamsAddIn.Location = New-Object System.Drawing.Point($column1X, 60)
$linkMissingTeamsAddIn.Add_LinkClicked({
        # Prompt the user for confirmation
        $result = [System.Windows.Forms.MessageBox]::Show("This action will force quit Outlook and Teams please save any important work beofre proceeding. Do you want to continue?", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            # Check if Teams and Outlook are running
            $teamsRunning = Get-Process -Name Teams -ErrorAction SilentlyContinue
            $outlookRunning = Get-Process -Name Outlook -ErrorAction SilentlyContinue

            # Terminate Teams and Outlook
            if ($teamsRunning) {
                Stop-Process -Name Teams -Force
                Write-Host "Teams has been terminated." -ForegroundColor Green
            }
            else {
                Write-Host "Teams was not running." -ForegroundColor Yellow
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
            if ($teamsRunning) {
                Start-Process -Name Teams
                Write-Host "Teams has been restarted." -ForegroundColor Green
            }

            if ($outlookRunning) {
                Start-Process -Name Outlook
                Write-Host "Outlook has been restarted." -ForegroundColor Green
            }

            [System.Windows.Forms.MessageBox]::Show("The process is complete. Teams and Outlook have been restarted if they were running before.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })
$sectionOfficeApps.Controls.Add($linkMissingTeamsAddIn)

# Create a hyperlink to Remove Office
$linkRemoveOffice = New-Object System.Windows.Forms.LinkLabel
$linkRemoveOffice.Text = "Remove Office"
$linkRemoveOffice.AutoSize = $true
$linkRemoveOffice.Location = New-Object System.Drawing.Point($column2X, 30)
$linkRemoveOffice.Add_LinkClicked({
        # Inform the user that the removal process is starting
        $result = [System.Windows.Forms.MessageBox]::Show("This action will close all running Office apps and remove all Office instances. Please save any important work before proceeding. Do you want to continue?", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            Write-Host "Removing Office..." -ForegroundColor Green

            # Step 1: Uninstall Office using the Office Removal Tool
            try {
                $officeRemovalToolPath = "$env:TEMP\OfficeRemovalTool.exe"
                Invoke-WebRequest -Uri "https://aka.ms/SaRA-officeUninstallFromPC" -OutFile $officeRemovalToolPath
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
                try {
                    Remove-Item -Recurse -Force -Path $folder
                    Write-Host "Removed folder: $folder" -ForegroundColor Green
                }
                catch {
                    Write-Host "Failed to remove folder: $folder" -ForegroundColor Red
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
                try {
                    Remove-Item -Recurse -Force -Path $regPath
                    Write-Host "Removed registry path: $regPath" -ForegroundColor Green
                }
                catch {
                    Write-Host "Failed to remove registry path: $regPath" -ForegroundColor Red
                }
            }

            # Step 4: Remove Office shortcuts
            $officeShortcuts = @(
                "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Office",
                "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office"
            )
            foreach ($shortcut in $officeShortcuts) {
                try {
                    Remove-Item -Recurse -Force -Path $shortcut
                    Write-Host "Removed shortcut: $shortcut" -ForegroundColor Green
                }
                catch {
                    Write-Host "Failed to remove shortcut: $shortcut" -ForegroundColor Red
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
                try {
                    Remove-Item -Recurse -Force -Path $tempFile
                    Write-Host "Removed temp file: $tempFile" -ForegroundColor Green
                }
                catch {
                    Write-Host "Failed to remove temp file: $tempFile" -ForegroundColor Red
                }
            }

            Write-Host "Office removal process is complete." -ForegroundColor Green

            # Prompt the user to reboot the computer
            [System.Windows.Forms.MessageBox]::Show("The Office removal process is complete. Please reboot your computer to finalize the changes.", "Reboot Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })
$sectionOfficeApps.Controls.Add($linkRemoveOffice)

# ...

# Run the form
$form.ShowDialog()
