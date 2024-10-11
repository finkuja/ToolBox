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

# Determine the script directory
if ($PSScriptRoot) {
    $scriptDir = $PSScriptRoot
}
elseif ($MyInvocation.MyCommand.Path) {
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
}
else {
    $scriptDir = [System.IO.Path]::GetDirectoryName([System.Reflection.Assembly]::GetExecutingAssembly().Location)
}

# Define the paths to the XAML and Functions folders
$xamlDir = [System.IO.Path]::Combine($scriptDir, "XAML")
$functionsDir = [System.IO.Path]::Combine($scriptDir, "Functions")

# Check if the XAML and Functions folders exist
$xamlExists = Test-Path -Path $xamlDir
$functionsExists = Test-Path -Path $functionsDir

if (-not $xamlExists -or -not $functionsExists) {
    Write-Host "XAML or Functions folder is missing." -ForegroundColor Red
    if (-not $OfflineMode) {
        # Define the URLs for the XAML and Functions folders
        $xamlUrl = "https://github.com/finkuja/ToolBox/refs/heads/Test/New%20WPF/XAML"
        $functionsUrl = "https://github.com/finkuja/ToolBox/refs/heads/Test/New%20WPF/Functions"

        # Define the paths to the temporary directories
        $tempDir = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "ESSToolBox")
        if (-not (Test-Path -Path $tempDir)) {
            New-Item -ItemType Directory -Path $tempDir | Out-Null
        }

        $xamlDir = [System.IO.Path]::Combine($tempDir, "XAML")
        $functionsDir = [System.IO.Path]::Combine($tempDir, "Functions")

        # Download the XAML and Functions folders
        Invoke-RestMethod -Uri "$xamlUrl/MainWindow.xml" -OutFile [System.IO.Path]::Combine($xamlDir, "MainWindow.xml")
        Invoke-RestMethod -Uri "$functionsUrl/Install.ps1" -OutFile [System.IO.Path]::Combine($functionsDir, "Install.ps1")
        Invoke-RestMethod -Uri "$functionsUrl/Tweak.ps1" -OutFile [System.IO.Path]::Combine($functionsDir, "Tweak.ps1")
        Invoke-RestMethod -Uri "$functionsUrl/Fix.ps1" -OutFile [System.IO.Path]::Combine($functionsDir, "Fix.ps1")
    }
    else {
        Write-Host "The XAML and Functions folders are missing and OfflineMode is enabled. Please ensure the folders are present." -ForegroundColor Red
        exit
    }
}

# Check if the XAML folder and MainWindow.xml file exist
$mainWindowPath = [System.IO.Path]::Combine($xamlDir, "MainWindow.xml")
$mainWindowExists = Test-Path -Path $mainWindowPath

if (-not $mainWindowExists) {
    Write-Host "The XAML folder or MainWindow.xml file cannot be found." -ForegroundColor Red
    exit
}

# Check if the Functions folder and required .ps1 files exist
$requiredFunctions = @("Install.ps1", "Tweak.ps1", "Fix.ps1")
foreach ($function in $requiredFunctions) {
    $functionPath = [System.IO.Path]::Combine($functionsDir, $function)
    $functionExists = Test-Path -Path $functionPath
    if (-not $functionExists) {
        Write-Host "The Functions folder or $function file cannot be found." -ForegroundColor Red
        exit
    }
}

# Load the PresentationFramework assembly for XAML support
Add-Type -AssemblyName PresentationFramework
# Load the winforms assembly for message box support
Add-Type -AssemblyName System.Windows.Forms

# Read the XAML file content
$xamlFilePath = [System.IO.Path]::Combine($xamlDir, "MainWindow.xml")
$xaml = Get-Content -Path $xamlFilePath -Raw

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
        try {
            # Clean up the temporary directory and its contents
            if (Test-Path -Path $tempDir) {
                Remove-Item -Path $tempDir -Recurse -Force
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
    })

# Find the MainTabControl
$mainTabControl = $window.FindName("MainTabControl")
if ($null -eq $mainTabControl) {
    Write-Host "MainTabControl not found in XAML." -ForegroundColor Red
    exit
}

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
        switch ($selectedTab.Name) {
            "InstallTab" {
                $installScriptPath = [System.IO.Path]::Combine($functionsDir, "Install.ps1")
                if (Test-Path -Path $installScriptPath) {
                    $installScriptContent = Get-Content -Path $installScriptPath -Raw
                    if ($installScriptContent) {
                        Invoke-Expression -Command $installScriptContent
                    }
                    else {
                        Write-Host "Install script is empty: $installScriptPath" -ForegroundColor Red
                    }
                }
                else {
                    Write-Host "Install script not found at path: $installScriptPath" -ForegroundColor Red
                }
            }
            "TweakTab" {
                $tweakScriptPath = [System.IO.Path]::Combine($functionsDir, "Tweak.ps1")
                if (Test-Path -Path $tweakScriptPath) {
                    $tweakScriptContent = Get-Content -Path $tweakScriptPath -Raw
                    if ($tweakScriptContent) {
                        Invoke-Expression -Command $tweakScriptContent
                    }
                    else {
                        Write-Host "Tweak script is empty: $tweakScriptPath" -ForegroundColor Red
                    }
                }
                else {
                    Write-Host "Tweak script not found at path: $tweakScriptPath" -ForegroundColor Red
                }
            }
            "FixTab" {
                $fixScriptPath = [System.IO.Path]::Combine($functionsDir, "Fix.ps1")
                if (Test-Path -Path $fixScriptPath) {
                    $fixScriptContent = Get-Content -Path $fixScriptPath -Raw
                    if ($fixScriptContent) {
                        Invoke-Expression -Command $fixScriptContent
                    }
                    else {
                        Write-Host "Fix script is empty: $fixScriptPath" -ForegroundColor Red
                    }
                }
                else {
                    Write-Host "Fix script not found at path: $fixScriptPath" -ForegroundColor Red
                }
            }
        }
    })

# Show the window
$window.ShowDialog() | Out-Null