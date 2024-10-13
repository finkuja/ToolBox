param (
    [switch]$OfflineMode
)

<#
.SYNOPSIS
    This script is the ESSToolBox, a tool for system administration tasks.

.DESCRIPTION
    The ESSToolBox is a PowerShell script designed to perform various system administration tasks. It includes features such as repairing Windows Update, checking for the installation of Windows Package Manager (winget), and more.

.PARAMETER OfflineMode
    If specified, the script will run in offline mode and will not download the modules. It will assume the dependencies are in the same folder as the script.

.NOTES
    - This script must be run with administrator privileges.
    - The script checks if Windows Package Manager (winget) is installed and prompts the user to install it if necessary.
    - The script also checks for the 'msstore' source in the winget source list and accepts the source agreements if necessary.

.DEPENDENCIES
    - Windows Package Manager (winget)
    - Microsoft.WinGet.Client module

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
 
 === Version Beta 0.2 ===

 === Author: Carlos Alvarez Magariños ===
 "

#########################
# Script Initialization #
#########################

# Check if the script is running with administrator privileges
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    Write-Host "The ESSToolBox needs to run as Administrator, trying to elevate the permissions..." -ForegroundColor Yellow
    
    # Get the current script content
    $scriptContent = (Invoke-RestMethod https://raw.githubusercontent.com/finkuja/ToolBox/refs/heads/main/EssTool.ps1)
    
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

# Check the system architecture
$cpuInfo = Get-WmiObject -Class Win32_Processor | Select-Object -First 1 -Property Name, Manufacturer, Description, Architecture

switch ($cpuInfo.Architecture) {
    0 {
        Write-Host "CPU Architecture: x86 (32-bit)"
        $disableInstall = $false
    }
    9 {
        Write-Host "CPU Architecture: x64 (64-bit)"
        $disableInstall = $false
    }
    5 {
        Write-Host "CPU Architecture: ARM"
        $disableInstall = $true
    }
    default {
        Write-Host "CPU Architecture: Unknown"
        $disableInstall = $true
    }
}

Write-Host "CPU Information: $($cpuInfo.Name), $($cpuInfo.Manufacturer), $($cpuInfo.Description)"

# Check if winget is installed
Write-Host "Checking if Windows Package Manager (winget) is installed..."
$winget = Get-Command winget -ErrorAction SilentlyContinue
if ($null -eq $winget) {
    [System.Windows.Forms.MessageBox]::Show("Windows Package Manager (winget) is not installed. Please install it from https://github.com/microsoft/winget-cli/releases")
    Start-Process "https://github.com/microsoft/winget-cli/releases"
    exit
}
else {
    Write-Host "Windows Package Manager (winget) is installed." -ForegroundColor Green
    
    # Update winget sources
    Write-Host "Updating winget sources..."
    try {
        Start-Process "winget" -ArgumentList "source update" -NoNewWindow -Wait
        Write-Host "Winget sources updated successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to update winget sources." -ForegroundColor Red
    }
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

# Import the winget module
try {
    Import-Module -Name Microsoft.WinGet.Client -ErrorAction Stop
    Write-Host "Microsoft.WinGet.Client module is imported." -ForegroundColor Green
}
catch {
    Write-Warning "Failed to import the Microsoft.WinGet.Client module. The Install Tab will not work properly."
}

# Check if PowerShellForGitHub module is installed and up to date, if not install or update it
$githubModuleName = "PowerShellForGitHub"
$githubModule = Get-Module -Name $githubModuleName -ListAvailable -ErrorAction SilentlyContinue

if (-not $githubModule) {
    Write-Host "Installing the $githubModuleName module..." -ForegroundColor Yellow
    try {
        Install-Module -Name $githubModuleName -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
        Write-Host "$githubModuleName module installed successfully." -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to install the $githubModuleName module"
    }
}
else {
    Write-Host "$githubModuleName module is already installed. Checking for updates..." -ForegroundColor Yellow
    try {
        $githubUpdateAvailable = Find-Module -Name $githubModuleName | Where-Object { $_.Version -gt $githubModule.Version }
        if ($githubUpdateAvailable) {
            Write-Host "Updating the $githubModuleName module to the latest version..." -ForegroundColor Yellow
            Update-Module -Name $githubModuleName -Force -ErrorAction Stop
            Write-Host "$githubModuleName module updated successfully." -ForegroundColor Green
        }
        else {
            Write-Host "$githubModuleName module is up to date." -ForegroundColor Green
        }
    }
    catch {
        Write-Warning "Failed to check for updates or update the $githubModuleName module."
    }
}

# Import the GitHub module
try {
    Import-Module -Name PowerShellForGitHub -ErrorAction Stop
    Write-Host "PowerShellForGitHub module is imported." -ForegroundColor Green
}
catch {
    Write-Warning "Failed to import the PowerShellForGitHub module. The Script will not function properly."
}

##################################################
# END of Check Install Update and Import Modules #
##################################################

# Initialize the tempDir flag
$tempDir = $false

# Determine the script directory
if ($PSScriptRoot) {
    $scriptDir = $PSScriptRoot
}
elseif ($MyInvocation.MyCommand.Path) {
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
}
else {
    # Use a temporary directory if the script directory cannot be determined
    $scriptDir = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "ESSToolBox")
    if (-not (Test-Path -Path $scriptDir)) {
        New-Item -ItemType Directory -Path $scriptDir | Out-Null
    }
    # Set the tempDir flag to true
    $tempDir = $true
}

Read-Host "Stop 0"
# Define the paths to the XAML and Functions folders
$xamlDir = [System.IO.Path]::Combine($scriptDir, "XAML")
Read-Host "Stop 1"
$functionsDir = [System.IO.Path]::Combine($scriptDir, "Functions")

Read-Host "Stop 2"

if ($tempDir) {
    Write-Host "Running from a temporary directory. Attempting to download XAML and Functions from GitHub..." -ForegroundColor Yellow
    Read-Host "Stop 4"
    # Define the GitHub repository and branch
    $repoOwner = "finkuja"
    $repoName = "ToolBox"
    $branch = "Test"
    Read-Host "Stop 5"
    # Define the paths to the XAML and Functions folders in the GitHub repository
    $xamlPath = "New WPF/XAML/MainWindow.xml"
    $functionsPath = "New WPF/Functions"
    Read-Host "Stop 0"
    # Define the local paths to the XAML and Functions folders
    $xamlDir = [System.IO.Path]::Combine($scriptDir, "XAML")
    Read-Host "Stop 1"
    $functionsDir = [System.IO.Path]::Combine($scriptDir, "Functions")
    
    # Create the local XAML and Functions directories if they do not exist
    if (-not (Test-Path -Path $xamlDir)) {
        New-Item -ItemType Directory -Path $xamlDir | Out-Null
    }
    if (-not (Test-Path -Path $functionsDir)) {
        New-Item -ItemType Directory -Path $functionsDir | Out-Null
    }
    
    # Download the MainWindow.xml file from the GitHub repository
    $mainWindowPath = [System.IO.Path]::Combine($xamlDir, "MainWindow.xml")
    try {
        Get-GitHubContent -RepositoryName $repoName -OwnerName $repoOwner -Ref $branch -Path $xamlPath -OutFile $mainWindowPath
        Write-Host "Downloaded MainWindow.xml successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to download MainWindow.xml: $_" -ForegroundColor Red
        exit
    }
    
    # Download all .ps1 files from the Functions folder in the GitHub repository
    try {
        $functionsContent = Get-GitHubContent -RepositoryName $repoName -OwnerName $repoOwner -Ref $branch -Path $functionsPath
        foreach ($file in $functionsContent) {
            if ($file.name -like "*.ps1") {
                $filePath = [System.IO.Path]::Combine($functionsDir, $file.name)
                Get-GitHubContent -RepositoryName $repoName -OwnerName $repoOwner -Ref $branch -Path "$functionsPath/$($file.name)" -OutFile $filePath
                Write-Host "Downloaded $($file.name) successfully." -ForegroundColor Green
            }
        }
    }
    catch {
        Write-Host "Failed to download Functions files: $_" -ForegroundColor Red
        exit
    }
}
else {
    Write-Host "Running from a local directory." -ForegroundColor Yellow
}

# Check if the MainWindow.xml file exists
$mainWindowPath = [System.IO.Path]::Combine($xamlDir, "MainWindow.xml")
if (-not (Test-Path -Path $mainWindowPath)) {
    Write-Host "The XAML folder or MainWindow.xml file cannot be found." -ForegroundColor Red
    exit
}

# Source all .ps1 files in the Functions directory
Get-ChildItem -Path $functionsDir -Filter *.ps1 | ForEach-Object {
    . $_.FullName
}

# Load the PresentationFramework assembly for XAML support
Add-Type -AssemblyName PresentationFramework
# Load the winforms assembly for message box support
Add-Type -AssemblyName System.Windows.Forms

################################
# END of Script Initialization #
################################

#######################################
#GUI Initialization and Event Handlers#
#######################################

# Load the XAML file content
$xaml = Get-Content -Path $mainWindowPath -Raw

# Load the XAML directly using XamlReader
try {
    $reader = (New-Object System.Xml.XmlTextReader (New-Object System.IO.StringReader $xaml))
    $window = [Windows.Markup.XamlReader]::Load($reader)
}
catch {
    Write-Host "Failed to load XAML: $_" -ForegroundColor Red
    exit
}

# Add the MouseLeftButtonDown event handler to make the window draggable
$windowControlPanel = $window.FindName("WindowControlPanel")
if ($null -eq $windowControlPanel) {
    Write-Host "WindowControlPanel not found in XAML." -ForegroundColor Red
    exit
}

# Load the XAML directly using XamlReader
try {
    $reader = (New-Object System.Xml.XmlTextReader (New-Object System.IO.StringReader $xaml))
    $window = [Windows.Markup.XamlReader]::Load($reader)
}
catch {
    Write-Host "Failed to load XAML: $_" -ForegroundColor Red
    exit
}

# Add the MouseLeftButtonDown event handler to make the window draggable
$windowControlPanel = $window.FindName("WindowControlPanel")
if ($null -eq $windowControlPanel) {
    Write-Host "WindowControlPanel not found in XAML." -ForegroundColor Red
    exit
}
$windowControlPanel.Add_MouseLeftButtonDown({
        param ($source, $e)
        $window.DragMove()
    })

# Find the CloseButton and add a Click event handler
$closeButton = $window.FindName("CloseButton")
if ($null -eq $closeButton) {
    Write-Host "CloseButton not found in XAML." -ForegroundColor Red
    exit
}
$closeButton.Add_Click({
        if ($tempDir) {
            try {
                # Clean up the temporary directory and its contents
                if (Test-Path -Path $scriptDir) {
                    Remove-Item -Path $scriptDir  -Recurse -Force
                }
            }
            catch {
                # Handle any errors that occur during the removal
                Write-Host "An error occurred while cleaning up the temporary directory: $_" -ForegroundColor Red
            }
            finally {
                # Close the window
                $window.Close()
            }
        }
        else {
            # Close the window
            $window.Close()
        }
    })

# Find the MainTabControl
$mainTabControl = $window.FindName("MainTabControl")
if ($null -eq $mainTabControl) {
    Write-Host "MainTabControl not found in XAML." -ForegroundColor Red
    exit
}


###########################################
# INSTALL TAB Event Handlers and Functions#
###########################################

# Hide the InstallTab if $disableInstall is $true
if ($disableInstall) {
    $installTab = $mainTabControl.Items | Where-Object { $_.Name -eq "InstallTab" }
    if ($installTab) {
        $installTab.Visibility = [System.Windows.Visibility]::Collapsed
    }

    # Set the default selected tab to Tweak
    $tweakTab = $mainTabControl.Items | Where-Object { $_.Name -eq "TweakTab" }
    if ($tweakTab) {
        $mainTabControl.SelectedItem = $tweakTab

        # Refresh the content of the TweakTab
        $tweakTabContent = $tweakTab.Content
        $tweakTab.Content = $null
        $tweakTab.Content = $tweakTabContent
    }
}

# Event handler for tab selection change
$mainTabControl.Add_SelectionChanged({
        param ($source, $e)
        $selectedTab = $source.SelectedItem
        if ($selectedTab) {
            $selectedTabName = $selectedTab.Name
            switch ($selectedTabName) {
                "InstallTab" {
                    # Refresh the content of the InstallTab
                    $installTabContent = $selectedTab.Content
                    $selectedTab.Content = $null
                    $selectedTab.Content = $installTabContent
                }
                "TweakTab" {
                    # Refresh the content of the TweakTab
                    $tweakTabContent = $selectedTab.Content
                    $selectedTab.Content = $null
                    $selectedTab.Content = $tweakTabContent
                }
                "FixTab" {
                    # Refresh the content of the FixTab
                    $fixTabContent = $selectedTab.Content
                    $selectedTab.Content = $null
                    $selectedTab.Content = $fixTabContent
                }
            }
        }
    })

# Find all checkboxes in the InstallTab
$checkboxes = @(
    $window.FindName("AdobeCreativeCloud"),
    $window.FindName("AdobeReaderDC"),
    $window.FindName("GoogleChrome"),
    $window.FindName("Fiddler"),
    $window.FindName("HWMonitor"),
    $window.FindName("DotNetAllVersions"),
    $window.FindName("MicrosoftEdge"),
    $window.FindName("MicrosoftOffice365"),
    $window.FindName("MicrosoftOneDrive"),
    $window.FindName("MicrosoftOneNote"),
    $window.FindName("MicrosoftTeams"),
    $window.FindName("MozillaFirefox"),
    $window.FindName("PowerAutomate"),
    $window.FindName("PowerBIDesktop"),
    $window.FindName("PowerToys"),
    $window.FindName("QuickAssist"),
    $window.FindName("RemoteDesktop"),
    $window.FindName("SARATool"),
    $window.FindName("SurfaceDiagnosticToolkit"),
    $window.FindName("VisioViewer2016"),
    $window.FindName("VisualStudioCode"),
    $window.FindName("SevenZip")
)

# Find the CheckAllButton and add a Click event handler
$checkAllButton = $window.FindName("CheckAllButton")
if ($null -eq $checkAllButton) {
    Write-Host "CheckAllButton not found in XAML." -ForegroundColor Red
    exit
}

$checkAllButton.Add_Click({
        $allChecked = $checkboxes | ForEach-Object { $_.IsChecked } | Where-Object { $_ -eq $false } | Measure-Object | Select-Object -ExpandProperty Count

        if ($allChecked -eq 0) {
            # Uncheck all checkboxes
            foreach ($checkbox in $checkboxes) {
                $checkbox.IsChecked = $false
            }
            $checkAllButton.Content = "Check All"
        }
        else {
            # Check all checkboxes
            foreach ($checkbox in $checkboxes) {
                $checkbox.IsChecked = $true
            }
            $checkAllButton.Content = "Uncheck All"
        }
    })

# Find the InstallButton and add a Click event handler
$installButton = $window.FindName("InstallButton")
if ($null -eq $installButton) {
    Write-Host "InstallButton not found in XAML." -ForegroundColor Red
    exit
}

$installButton.Add_Click({
        # Get the names of the checked checkboxes
        $checkedItems = $checkboxes | Where-Object { $_.IsChecked -eq $true } | ForEach-Object { $_.Tag }

        # Invoke the Invoke-WinGet script with the checked items
        foreach ($item in $checkedItems) {
            Write-Host "Installing $item..."
            Invoke-WinGet -PackageName $item -Action Install -window $window
        }
    })
    
# Find the UninstallButton and add a Click event handler
$uninstallButton = $window.FindName("UninstallButton")
if ($null -eq $uninstallButton) {
    Write-Host "UninstallButton not found in XAML." -ForegroundColor Red
    exit
}
$uninstallButton.Add_Click({
        # Get the names of the checked checkboxes
        $checkedItems = $checkboxes | Where-Object { $_.IsChecked -eq $true } | ForEach-Object { $_.Tag }

        # Invoke the Invoke-WinGet script with the checked items
        foreach ($item in $checkedItems) {
            Write-Host "Uninstalling $item..."
            Invoke-WinGet -PackageName $item -Action Uninstall -window $window
        }
    })

# Find the InstalledButton and add a Click event handler
$installedButton = $window.FindName("InstalledButton")
if ($null -eq $installedButton) {
    Write-Host "InstalledButton not found in XAML." -ForegroundColor Red
    exit
}
$installedButton.Add_Click({
        # Check if the ShowAllInstalled checkbox is checked
        $showAllInstalledCheckbox = $window.FindName("ShowAllInstalled")
        if ($showAllInstalledCheckbox -and $showAllInstalledCheckbox.IsChecked) {
            # Run WinGet List and display the results in the console
            Start-Process "winget" -ArgumentList "list" -NoNewWindow -Wait
            return
        }

        # Run the Get-WinGetPackage command and capture the output directly
        $output = Get-WinGetPackage
        # Extract the package names from the output
        $packageNames = $output | Select-Object -ExpandProperty Name

        # Iterate through each checkbox in the $checkboxes array
        foreach ($checkbox in $checkboxes) {
            $checkboxName = $checkbox.Tag
            # Check if the checkbox name is present in the package names
            if ($packageNames -contains $checkboxName) {
                $checkbox.IsChecked = $true
            }
            else {
                $checkbox.IsChecked = $false
            }
        }
    })

###################################################
# END OF INSTALL TAB Event Handlers and Functions #
###################################################


##########################################
# TWEAK TAB Event Handlers and Functions #
##########################################

# Find all checkboxes in the TweakTab
$tweakCheckBoxes = @(
    $window.FindName("CleanBoot"),
    $window.FindName("EnableDetailedBSODInformation"),
    $window.FindName("EnableGodMode"),
    $window.FindName("EnableClassicRightClickMenu"),
    $window.FindName("EnableEndTaskWithRightClick"),
    $window.FindName("ChangeIRPStackSize"),
    $window.FindName("ClipboardHistory"),
    $window.FindName("EnableVerboseLogonMessages"),
    $window.FindName("EnableVerboseStartupAndShutdownMessages")
)

# Find the ApplyButton and add a Click event handler
$ApplyButton = $window.FindName("ApplyButton")
if ($null -eq $ApplyButton) {
    Write-Error "ApplyButton not found."
}
else {

    $ApplyButton.Add_Click({
            # Get the names of the checked checkboxes
            $checkedItems = $tweakCheckBoxes | Where-Object { $_.IsChecked -eq $true } | ForEach-Object { $_.Name }
            foreach ($item in $checkedItems) {
                Invoke-Tweak -Action "Apply" -window $window -Tweak $item
            }
        })
}

# Find the UndoButton and add a Click event handler
$UndoButton = $window.FindName("UndoButton")
if ($null -eq $UndoButton) {
    Write-Error "UndoButton not found."
}
else {
    $UndoButton.Add_Click({
            $checkedItems = $tweakCheckBoxes | Where-Object { $_.IsChecked -eq $true } | ForEach-Object { $_.Name }
            foreach ($item in $checkedItems) {
                Invoke-Tweak -Action "Undo" -window $window -Tweak $item
            }
        })
}

# Find the DeleteTempFilesButton and add a Click event handler
$DeleteTempFilesButton = $window.FindName("DeleteTempFilesButton")
if ($null -eq $DeleteTempFilesButton) {
    Write-Error "DeleteTempFilesButton not found."
}
else {
    $DeleteTempFilesButton.Add_Click({
            Invoke-DeleteTempFiles
        })
}

# Find the OptimizeDrivesButton and add a Click event handler
$OptimizeDrivesButton = $window.FindName("OptimizeDrivesButton")
if ($null -eq $OptimizeDrivesButton) {
    Write-Error "OptimizeDrivesButton not found."
}
else {
    $OptimizeDrivesButton.Add_Click({
            
            Invoke-OptimizeDrives
        })
}

# Find the RunDiskCleanupButton and add a Click event handler
$RunDiskCleanupButton = $window.FindName("RunDiskCleanupButton")
if ($null -eq $RunDiskCleanupButton) {
    Write-Error "RunDiskCleanupButton not found."
}
else {
    $RunDiskCleanupButton.Add_Click({
            Invoke-RunDiskCleanup
        })
}

# Find the DNSComboBox and add a SelectionChanged event handler
$DNSComboBox = $window.FindName("DNSComboBox")
if ($null -eq $DNSComboBox) {
    Write-Error "DNSComboBox not found."
}
else {
    $DNSComboBox.Add_SelectionChanged({
            # Add your logic for handling DNS selection change here
            $selectedDNS = $DNSComboBox.SelectedItem.Content
            Write-Host "DNS selection changed to: $selectedDNS"
        })
}


#################################################
# END OF TWEAK TAB Event Handlers and Functions #
#################################################

########################################
# FIX TAB Event Handlers and Functions #
########################################

###############################################
# END OF FIX TAB Event Handlers and Functions #
###############################################

# Show the window
$window.ShowDialog() | Out-Null