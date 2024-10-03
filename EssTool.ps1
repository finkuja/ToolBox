# Ensure you run this script with administrator privileges


Write-Output " ______  _____  _____    _______          _   ____ 
|  ____|/ ____|/ ____|  |__   __|        | | |  _ \             
| |__  | (___ | (___       | | ___   ___ | | | |_) | ___  __  __
|  __|  \___ \ \___ \      | |/ _ \ / _ \| | |  _ < / _ \ \ \/ /
| |____ ____) |____) |     | | (_) | (_) | | | |_) | (_) | |  |
|______|_____/|_____/      |_|\___/ \___/|_| |____/ \___/ /_/\_\
 
 === Version Beta 0.1 ===

 === Created by: Carlos Alvarez Magariños ===
 "
# Import necessary .NET assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Check if the script is running with administrator privileges
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    Write-Host "The ESSToolBox needs to run as Administrator, trying to elevate the permissions..." -ForegroundColor Yellow
    
    # Get the current script content
    $scriptContent = (irm aka.ms/esstoolbox)
    
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
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to run the script with administrator privileges. Please run this script with administrator privileges manually.")
    }
    
    # Exit the current script
    exit
} else {
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

# Get the winget source list and find the 'msstore' source
$wingetSource = & winget source list | Where-Object { $_.Name -eq 'msstore' }

# Check if the 'msstore' source exists and if the source agreement is not accepted
if ($wingetSource) {
    if (-not $wingetSource.Accepted) {
        # Update the 'msstore' source and accept the source agreements
        Start-Process "winget" -ArgumentList "source update --name msstore" -NoNewWindow -Wait
        Write-Host "MS Store Source Agreement has been accepted." -ForegroundColor Green
    } else {
        Write-Host "MS Store Source Agreement is already accepted." -ForegroundColor Green
    }
} else {
    Write-Host "MS Store source not found." -ForegroundColor Yellow
}


# ----------------- Declaration of Functions for Cemplex Fix and Tweak Scenarios------------------
#___________________________________________________________________________________________________


# Function from Chris Titus Tech winutils.ps1 script github.com/christitus/winutils
function Invoke-FixesWUpdate {

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
                ($percent = try {(
                    $_.Substring(
                        $index + 6,
                        $_.IndexOf("%", $index) - $index - 6
                    )
                ).Trim()} catch {0}) `
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
                    [int]$percent = try {(
                        (
                            $_.Substring(
                                $_.IndexOf("n") + 2,
                                $_.IndexOf("%") - $_.IndexOf("n") - 2
                            ).ToCharArray() | Where-Object {$_}
                        ) -join ''
                    ).TrimStart()} catch {0}
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
                } catch {0}) `
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
                    [int]$percent = try {(
                        (
                            $_.Substring(
                                $_.IndexOf("n") + 2,
                                $_.IndexOf("%") - $_.IndexOf("n") - 2
                            ).ToCharArray() | Where-Object {$_}
                        ) -join ''
                    ).TrimStart()} catch {0}
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

#----------------------------------------------------------------------------------------------------------------------------------
#__________________________________________________________________________________________________________________________________


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

# Create the Install tab
$tabInstall = New-Object System.Windows.Forms.TabPage
$tabInstall.Text = "Install / Update"
$tabControl.Controls.Add($tabInstall)

# Define Visual Elements for All Tabs
# Define column positions
$column1X = 20
$column2X = 200
#$column3X = 380
#Define Section Length
$sectionLength = 520
#Define Install and Tweak Tab Buttons y position
$buttonY = 340

# Create checkboxes for packages in the Install tab
$checkboxAdobe = New-Object System.Windows.Forms.CheckBox
$checkboxAdobe.Text = "Adobe Reader"
$checkboxAdobe.Name = "Adobe"
$checkboxAdobe.AutoSize = $true
$checkboxAdobe.Location = New-Object System.Drawing.Point($column1X, 20)
$tabInstall.Controls.Add($checkboxAdobe)

$checkboxAdobeCloud = New-Object System.Windows.Forms.CheckBox
$checkboxAdobeCloud.Text = "Adobe Creative Cloud"
$checkboxAdobeCloud.Name = "Adobe.Cloud"
$checkboxAdobeCloud.AutoSize = $true
$checkboxAdobeCloud.Location = New-Object System.Drawing.Point($column1X, 50)
$tabInstall.Controls.Add($checkboxAdobeCloud)

$checkboxOffice = New-Object System.Windows.Forms.CheckBox
$checkboxOffice.Text = "Microsoft Office 365"
$checkboxOffice.Name = "Microsoft.Office"
$checkboxOffice.AutoSize = $true
$checkboxOffice.Location = New-Object System.Drawing.Point($column1X, 80)
$tabInstall.Controls.Add($checkboxOffice)

$checkboxOneNote = New-Object System.Windows.Forms.CheckBox
$checkboxOneNote.Text = "Microsoft OneNote"
$checkboxOneNote.Name = "Microsoft.OneNote"
$checkboxOneNote.AutoSize = $true
$checkboxOneNote.Location = New-Object System.Drawing.Point($column1X, 110)
$tabInstall.Controls.Add($checkboxOneNote)

$checkboxTeams = New-Object System.Windows.Forms.CheckBox
$checkboxTeams.Text = "Microsoft Teams"
$checkboxTeams.Name = "Microsoft.Teams"
$checkboxTeams.AutoSize = $true
$checkboxTeams.Location = New-Object System.Drawing.Point($column1X, 140)
$tabInstall.Controls.Add($checkboxTeams)

$checkboxNetFrameworks = New-Object System.Windows.Forms.CheckBox
$checkboxNetFrameworks.Text = ".NET Frameworks"
$checkboxNetFrameworks.Name = "NetFrameworks"
$checkboxNetFrameworks.AutoSize = $true
$checkboxNetFrameworks.Location = New-Object System.Drawing.Point($column1X, 170)
$tabInstall.Controls.Add($checkboxNetFrameworks)

$checkboxPowerAutomate = New-Object System.Windows.Forms.CheckBox
$checkboxPowerAutomate.Text = "Power Automate"
$checkboxPowerAutomate.Name = "PowerAutomate"
$checkboxPowerAutomate.AutoSize = $true
$checkboxPowerAutomate.Location = New-Object System.Drawing.Point($column1X, 200)
$tabInstall.Controls.Add($checkboxPowerAutomate)

$checkboxPowerToys = New-Object System.Windows.Forms.CheckBox
$checkboxPowerToys.Text = "PowerToys"
$checkboxPowerToys.Name = "Microsoft.PowerToys"
$checkboxPowerToys.AutoSize = $true
$checkboxPowerToys.Location = New-Object System.Drawing.Point($column1X, 230)
$tabInstall.Controls.Add($checkboxPowerToys)

$checkboxQuickAssist = New-Object System.Windows.Forms.CheckBox
$checkboxQuickAssist.Text = "Quick Assist"
$checkboxQuickAssist.Name = "QuickAssist"
$checkboxQuickAssist.AutoSize = $true
$checkboxQuickAssist.Location = New-Object System.Drawing.Point($column1X, 260)
$tabInstall.Controls.Add($checkboxQuickAssist)

$checkboxSurfaceDiagnosticToolkit = New-Object System.Windows.Forms.CheckBox
$checkboxSurfaceDiagnosticToolkit.Text = "Surface Diagnostic Toolkit"
$checkboxSurfaceDiagnosticToolkit.Name = "SurfaceDiagnosticToolkit"
$checkboxSurfaceDiagnosticToolkit.AutoSize = $true
$checkboxSurfaceDiagnosticToolkit.Location = New-Object System.Drawing.Point($column2X, 20)
$tabInstall.Controls.Add($checkboxSurfaceDiagnosticToolkit)

$checkboxVisio = New-Object System.Windows.Forms.CheckBox
$checkboxVisio.Text = "Visio"
$checkboxVisio.Name = "Visio"
$checkboxVisio.AutoSize = $true
$checkboxVisio.Location = New-Object System.Drawing.Point($column2X, 50)
$tabInstall.Controls.Add($checkboxVisio)

# Create an Install button in the Install tab
$buttonInstall = New-Object System.Windows.Forms.Button
$buttonInstall.Text = "Install"
$buttonInstall.AutoSize = $true
$buttonInstall.Location = New-Object System.Drawing.Point(20, $buttonY)
$tabInstall.Controls.Add($buttonInstall)

# Define the action for the Install button
$buttonInstall.Add_Click({
    if ($checkboxOffice.Checked) {
        Start-Process "winget" -ArgumentList "install --id Microsoft.Office -e" -NoNewWindow -Wait
    }
    if ($checkboxPowerToys.Checked) {
        Start-Process "winget" -ArgumentList "install --id Microsoft.PowerToys -e" -NoNewWindow -Wait
    }
    if ($checkboxTeams.Checked) {
        Start-Process "winget" -ArgumentList "install --id Microsoft.Teams -e" -NoNewWindow -Wait
    }
    if ($checkboxOneNote.Checked) {
        Start-Process "winget" -ArgumentList "install --id Microsoft.OneNote -e" -NoNewWindow -Wait
    }
    if ($checkboxOffice.Checked) {
        Start-Process "winget" -ArgumentList "install --id Microsoft.Office -e" -NoNewWindow -Wait
    }
    if ($checkboxPowerToys.Checked) {
        Start-Process "winget" -ArgumentList "install --id Microsoft.PowerToys -e" -NoNewWindow -Wait
    }
    if ($checkboxTeams.Checked) {
        Start-Process "winget" -ArgumentList "install --id Microsoft.Teams -e" -NoNewWindow -Wait
    }
    if ($checkboxOneNote.Checked) {
        Start-Process "winget" -ArgumentList "install --id Microsoft.OneNote -e" -NoNewWindow -Wait
    }
    if ($checkboxAdobe.Checked) {
        Start-Process "winget" -ArgumentList "install --id Adobe -e" -NoNewWindow -Wait
    }
    if ($checkboxAdobeCloud.Checked) {
        Start-Process "winget" -ArgumentList "install --id Adobe.Cloud -e" -NoNewWindow -Wait
    }
    if ($checkboxPowerAutomate.Checked) {
        Start-Process "winget" -ArgumentList "install --id PowerAutomate -e" -NoNewWindow -Wait
    }
    if ($checkboxVisio.Checked) {
        Start-Process "winget" -ArgumentList "install --id Visio -e" -NoNewWindow -Wait
    }
    if ($checkboxNetFrameworks.Checked) {
        Start-Process "winget" -ArgumentList "install --id NetFrameworks -e" -NoNewWindow -Wait
    }
    if ($checkboxQuickAssist.Checked) {
        Start-Process "winget" -ArgumentList "install --id QuickAssist -e" -NoNewWindow -Wait
    }
    if ($checkboxSurfaceDiagnosticToolkit.Checked) {
        Start-Process "winget" -ArgumentList "install --id SurfaceDiagnosticToolkit -e" -NoNewWindow -Wait
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
    if ($checkboxOffice.Checked) {
        Start-Process "winget" -ArgumentList "uninstall --id Microsoft.Office -e" -NoNewWindow -Wait
    }
    if ($checkboxPowerToys.Checked) {
        Start-Process "winget" -ArgumentList "uninstall --id Microsoft.PowerToys -e" -NoNewWindow -Wait
    }
    if ($checkboxTeams.Checked) {
        Start-Process "winget" -ArgumentList "uninstall --id Microsoft.Teams -e" -NoNewWindow -Wait
    }
    if ($checkboxOneNote.Checked) {
        Start-Process "winget" -ArgumentList "uninstall --id Microsoft.OneNote -e" -NoNewWindow -Wait
    }
    if ($checkboxAdobe.Checked) {
        Start-Process "winget" -ArgumentList "uninstall --id Adobe -e" -NoNewWindow -Wait
    }
    if ($checkboxAdobeCloud.Checked) {
        Start-Process "winget" -ArgumentList "uninstall --id Adobe.Cloud -e" -NoNewWindow -Wait
    }
    if ($checkboxPowerAutomate.Checked) {
        Start-Process "winget" -ArgumentList "uninstall --id PowerAutomate -e" -NoNewWindow -Wait
    }
    if ($checkboxVisio.Checked) {
        Start-Process "winget" -ArgumentList "uninstall --id Visio -e" -NoNewWindow -Wait
    }
    if ($checkboxNetFrameworks.Checked) {
        Start-Process "winget" -ArgumentList "uninstall --id NetFrameworks -e" -NoNewWindow -Wait
    }
    if ($checkboxQuickAssist.Checked) {
        Start-Process "winget" -ArgumentList "uninstall --id QuickAssist -e" -NoNewWindow -Wait
    }
    if ($checkboxSurfaceDiagnosticToolkit.Checked) {
        Start-Process "winget" -ArgumentList "uninstall --id SurfaceDiagnosticToolkit -e" -NoNewWindow -Wait
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
            } else {
                $control.Checked = $false
            }
        }
    }
})

# Create the Tweak tab
$tabTweak = New-Object System.Windows.Forms.TabPage
$tabTweak.Text = "Tweak"
$tabControl.Controls.Add($tabTweak)

# Create controls for the Tweak tab
$checkboxRightClickEndTask = New-Object System.Windows.Forms.CheckBox
$checkboxRightClickEndTask.Text = "Enable end task with right click"
$checkboxRightClickEndTask.Name = "EnableRightClickEndTask"
$checkboxRightClickEndTask.AutoSize = $true
$checkboxRightClickEndTask.Location = New-Object System.Drawing.Point($column1X, 80)
$tabTweak.Controls.Add($checkboxRightClickEndTask)

$checkboxRunDiskCleanup = New-Object System.Windows.Forms.CheckBox
$checkboxRunDiskCleanup.Text = "Run disk cleanup"
$checkboxRunDiskCleanup.Name = "RunDiskCleanup"
$checkboxRunDiskCleanup.AutoSize = $true
$checkboxRunDiskCleanup.Location = New-Object System.Drawing.Point($column1X, 140)
$tabTweak.Controls.Add($checkboxRunDiskCleanup)

$checkboxDetailedBSOD = New-Object System.Windows.Forms.CheckBox
$checkboxDetailedBSOD.Text = "Enable detailed BSOD information"
$checkboxDetailedBSOD.Name = "EnableDetailedBSOD"
$checkboxDetailedBSOD.AutoSize = $true
$checkboxDetailedBSOD.Location = New-Object System.Drawing.Point($column1X, 50)
$tabTweak.Controls.Add($checkboxDetailedBSOD)

$checkboxVerboseLogon = New-Object System.Windows.Forms.CheckBox
$checkboxVerboseLogon.Text = "Enable verbose logon messages"
$checkboxVerboseLogon.Name = "EnableVerboseLogon"
$checkboxVerboseLogon.AutoSize = $true
$checkboxVerboseLogon.Location = New-Object System.Drawing.Point($column1X, 110)
$tabTweak.Controls.Add($checkboxVerboseLogon)

$checkboxDeleteTempFiles = New-Object System.Windows.Forms.CheckBox
$checkboxDeleteTempFiles.Text = "Delete temporary files"
$checkboxDeleteTempFiles.Name = "DeleteTempFiles"
$checkboxDeleteTempFiles.AutoSize = $true
$checkboxDeleteTempFiles.Location = New-Object System.Drawing.Point($column1X, 20)
$tabTweak.Controls.Add($checkboxDeleteTempFiles)

$buttonApply = New-Object System.Windows.Forms.Button
$buttonApply.Text = "Apply"
$buttonApply.Location = New-Object System.Drawing.Point(20, $buttonY)
$tabTweak.Controls.Add($buttonApply)

$buttonUndo = New-Object System.Windows.Forms.Button
$buttonUndo.Text = "Undo"
$buttonUndo.Location = New-Object System.Drawing.Point(120, $buttonY)
$tabTweak.Controls.Add($buttonUndo)

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
        } catch {
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
        } catch {
            Write-Host "Failed to enable verbose logon messages: $_" -ForegroundColor Red
        }
    }
    if ($checkboxDeleteTempFiles.Checked) {
        # Delete temporary files
        Write-Host "Deleting temporary files..." -ForegroundColor Green

        # Check if the C:\Windows\Temp path exists
        if (Test-Path "C:\Windows\Temp") {
            Get-ChildItem -Path "C:\Windows\Temp" *.* -Recurse | Remove-Item -Force -Recurse
        } else {
            Write-Host "Path C:\Windows\Temp does not exist." -ForegroundColor Red
        }

        # Check if the $env:TEMP path exists
        if (Test-Path $env:TEMP) {
            Get-ChildItem -Path $env:TEMP *.* -Recurse | Remove-Item -Force -Recurse
        } else {
            Write-Host "Path $env:TEMP does not exist." -ForegroundColor Red
        }

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
            } else {
                Write-Host "Registry path does not exist: $registryPath" -ForegroundColor Red
            }
        } catch {
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
            } else {
                Write-Host "Registry path does not exist: $registryPath" -ForegroundColor Red
            }
        } catch {
            Write-Host "Failed to disable verbose logon messages: $_" -ForegroundColor Red
        }
    }
    if ($checkboxDeleteTempFiles.Checked) {
        # Undo delete temporary files
        Write-Output "Nothing to do here..."
        # Add your code here
    }
    [System.Windows.Forms.MessageBox]::Show("Selected tweaks have been undone.")
})

# Create the Fix tab
$tabFix = New-Object System.Windows.Forms.TabPage
$tabFix.Text = "Fix"
$tabControl.Controls.Add($tabFix)

# Create controls for the Fix tab

# Apps section
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
    [string]$url="https://swupmf.adobe.com/webfeed/CleanerTool/win/AdobeCreativeCloudCleanerTool.exe"

    Write-Host "The Adobe Creative Cloud Cleaner tool is hosted at"
    Write-Host "$url"

    try {
        # Don't show the progress because it will slow down the download speed
        $ProgressPreference='SilentlyContinue'

        Invoke-WebRequest -Uri $url -OutFile "$env:TEMP\AdobeCreativeCloudCleanerTool.exe" -UseBasicParsing -ErrorAction SilentlyContinue -Verbose

        # Revert back the ProgressPreference variable to the default value since we got the file desired
        $ProgressPreference='Continue'

        Start-Process -FilePath "$env:TEMP\AdobeCreativeCloudCleanerTool.exe" -Wait -ErrorAction SilentlyContinue -Verbose
    } catch {
        Write-Error $_.Exception.Message
    } finally {
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
    # Add your code here
})
$sectionApps.Controls.Add($linkRemoveAdobeReader)

# System section
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
    $wifiAdapter = Get-NetAdapter | Where-Object {$_.InterfaceDescription -like "*Wi-Fi*"}
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

# Office Apps section
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
    # Rebuild .OST file
    Write-Output "Rebuilding .OST file..."
    # Add your code here
})
$sectionOfficeApps.Controls.Add($linkRebuildOST)

# Create a hyperlink for Missing Teams Add-In
$linkMissingTeamsAddIn = New-Object System.Windows.Forms.LinkLabel
$linkMissingTeamsAddIn.Text = "Missing Teams Add-In"
$linkMissingTeamsAddIn.AutoSize = $true
$linkMissingTeamsAddIn.Location = New-Object System.Drawing.Point($column1X, 60)
$linkMissingTeamsAddIn.Add_LinkClicked({
    # Fix Missing Teams Add-In
    Write-Output "Fixing Missing Teams Add-In..."
    # Add your code here
})
$sectionOfficeApps.Controls.Add($linkMissingTeamsAddIn)

# Create a hyperlink to Remove Office
$linkRemoveOffice = New-Object System.Windows.Forms.LinkLabel
$linkRemoveOffice.Text = "Remove Office"
$linkRemoveOffice.AutoSize = $true
$linkRemoveOffice.Location = New-Object System.Drawing.Point($column2X, 30)
$linkRemoveOffice.Add_LinkClicked({
    # Remove Office
    Write-Output "Removing Office..."
    # Add your code here
})
$sectionOfficeApps.Controls.Add($linkRemoveOffice)


# ...

# Run the form
$form.ShowDialog()
