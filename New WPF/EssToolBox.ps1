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

# Determine the script directory
if ($PSScriptRoot) {
    $scriptDir = $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
} else {
    $scriptDir = [System.IO.Path]::GetDirectoryName([System.Reflection.Assembly]::GetExecutingAssembly().Location)
}

# Define the URLs for the module files
$installModuleUrl = "https://raw.githubusercontent.com/finkuja/ToolBox/main/Content/InstallGUI.psm1"
$tweakModuleUrl = "https://raw.githubusercontent.com/finkuja/ToolBox/main/Content/TweakGUI.psm1"
$fixModuleUrl = "https://raw.githubusercontent.com/finkuja/ToolBox/main/Content/FixGUI.psm1"

# Define the paths to the modules in the temporary directory
$tempDir = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "ESSToolBox")
if (-not (Test-Path -Path $tempDir)) {
    New-Item -ItemType Directory -Path $tempDir | Out-Null
}

$installModulePath = [System.IO.Path]::Combine($tempDir, "InstallGUI.psm1")
$tweakModulePath = [System.IO.Path]::Combine($tempDir, "TweakGUI.psm1")
$fixModulePath = [System.IO.Path]::Combine($tempDir, "FixGUI.psm1")

# Download the module files
Invoke-RestMethod -Uri $installModuleUrl -OutFile $installModulePath
Invoke-RestMethod -Uri $tweakModuleUrl -OutFile $tweakModulePath
Invoke-RestMethod -Uri $fixModuleUrl -OutFile $fixModulePath

# Check if the Install module exists and import it
if (Test-Path -Path $installModulePath) {
    Import-Module -Name $installModulePath -Force
} else {
    Write-Host "The Install module at path '$installModulePath' cannot be found." -ForegroundColor Red
}

# Check if the Tweak module exists and import it
if (Test-Path -Path $tweakModulePath) {
    Import-Module -Name $tweakModulePath -Force
} else {
    Write-Host "The Tweak module at path '$tweakModulePath' cannot be found." -ForegroundColor Red
}

# Check if the Fix module exists and import it
if (Test-Path -Path $fixModulePath) {
    Import-Module -Name $fixModulePath -Force
} else {
    Write-Host "The Fix module at path '$fixModulePath' cannot be found." -ForegroundColor Red
}

################################
# Main Menu Display  using WPF #
################################

# Load the XAML form
Add-Type -AssemblyName PresentationFramework

# Define the XAML for the main window with rounded corners using WindowChrome
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="ESSToolBox" Height="450" Width="800" WindowStartupLocation="CenterScreen" AllowsTransparency="True" WindowStyle="None" Background="Transparent" ResizeMode="NoResize">
    <WindowChrome.WindowChrome>
        <WindowChrome CornerRadius="10" GlassFrameThickness="0" UseAeroCaptionButtons="False"/>
    </WindowChrome.WindowChrome>
    <Border Background="#383131" CornerRadius="10" BorderBrush="Gray" BorderThickness="1">
        <Grid Name="WindowControlPanel">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <DockPanel LastChildFill="True">
                <Button Name="CloseButton" DockPanel.Dock="Right" Width="40" Height="40" Margin="5" Background="Transparent" BorderBrush="Transparent" Foreground="White" Cursor="Hand">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="0">
                                <TextBlock x:Name="CloseIcon" Text="&#xE10A;" FontFamily="Segoe MDL2 Assets" Foreground="{TemplateBinding Foreground}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter TargetName="CloseIcon" Property="Foreground" Value="#E30101"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Button.Template>
                </Button>
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="ESSToolBox" Margin="10,0,0,0" FontSize="16" FontWeight="Bold" Foreground="White"/>
                    <TextBlock Text="Beta v 0.2" Margin="5,0,0,0" FontSize="16" Foreground="White"/>
                </StackPanel>
            </DockPanel>
            <TabControl Grid.Row="1" Name="MainTabControl" Background="Transparent" BorderBrush="Transparent" BorderThickness="0">
                <TabControl.Resources>
                    <Style TargetType="TabItem">
                        <Setter Property="Template">
                            <Setter.Value>
                                <ControlTemplate TargetType="TabItem">
                                    <Border Name="Border" Background="Transparent" BorderBrush="Transparent" BorderThickness="0" CornerRadius="10" Margin="2">
                                        <ContentPresenter x:Name="ContentSite" VerticalAlignment="Center" HorizontalAlignment="Center" ContentSource="Header" Margin="10,2"/>
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsSelected" Value="True">
                                            <Setter TargetName="Border" Property="Background" Value="#185FF8"/>
                                            <Setter Property="Foreground" Value="White"/>
                                        </Trigger>
                                        <Trigger Property="IsSelected" Value="False">
                                            <Setter Property="Foreground" Value="#185FF8"/>
                                        </Trigger>
                                        <Trigger Property="IsEnabled" Value="False">
                                            <Setter Property="Foreground" Value="Gray"/>
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Setter.Value>
                        </Setter>
                    </Style>
                </TabControl.Resources>
                <TabItem Header="Install" Name="InstallTab" FontSize="14">
                    <Grid Name="InstallTabGrid" Background="Transparent">
                        <TextBlock Text="Install Content" VerticalAlignment="Top" HorizontalAlignment="Left" Margin="10"/>
                    </Grid>
                </TabItem>
                <TabItem Header="Tweak" Name="TweakTab" FontSize="14">
                    <Grid Background="Transparent">
                        <TextBlock Text="Tweak Content" VerticalAlignment="Top" HorizontalAlignment="Left" Margin="10"/>
                    </Grid>
                </TabItem>
                <TabItem Header="Fix" Name="FixTab" FontSize="14">
                    <Grid Background="Transparent">
                        <TextBlock Text="Fix Content" VerticalAlignment="Top" HorizontalAlignment="Left" Margin="10"/>
                    </Grid>
                </TabItem>
            </TabControl>
        </Grid>
    </Border>
</Window>
"@

# Load the XAML
try {
    [xml]$xamlWindow = $xaml
    $reader = (New-Object System.Xml.XmlNodeReader $xamlWindow)
    $window = [Windows.Markup.XamlReader]::Load($reader)
} catch {
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
    # Clean up the temporary directory and its contents
    if (Test-Path -Path $tempDir) {
        Remove-Item -Path $tempDir -Recurse -Force
    }
    $window.Close()
})

# Find the MainTabControl
$mainTabControl = $window.FindName("MainTabControl")
if ($null -eq $mainTabControl) {
    Write-Host "MainTabControl not found in XAML." -ForegroundColor Red
    exit
}

# Event handler for tab selection change
$mainTabControl.Add_SelectionChanged({
    param ($source, $e)
    $selectedTab = $source.SelectedItem
    switch ($selectedTab.Name) {
        "InstallTab" {
            Invoke-Install -DisableInstall:$disableInstall -MainWindow $window
        }
        "TweakTab" {
            # Add corresponding function call for TweakTab if needed
        }
        "FixTab" {
            # Add corresponding function call for FixTab if needed
        }
    }
})

# Show the window
$window.ShowDialog() | Out-Null