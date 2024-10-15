## Main Script

### ESSToolBox.ps1
This is the main script that initializes the GUI and handles the primary logic for the tool. It loads the main window from the `mainWindow.xaml` file and sets up event handlers for various buttons and tabs.

## Supporting Scripts

### Functions/
- **Show-ChildWindow.ps1**: Dynamically loads the contents of a child window based on the provided XAML and JSON files.
- **Invoke-WinGet.ps1**: Handles installation and uninstallation of packages using WinGet.
- **Invoke-Tweak.ps1**: Applies or undoes system tweaks.
- **Invoke-DeleteTempFiles.ps1**: Deletes temporary files.
- **Invoke-OptimizeDrives.ps1**: Optimizes drives.
- **Invoke-DiskCleanup.ps1**: Runs disk cleanup.
- **Reset-EdgeCache.ps1**: Resets Edge browser cache.
- **Reset-EdgeProfile.ps1**: Resets Edge browser profile.
- **Remove-Edge.ps1**: Removes Edge browser.
- **Invoke-FixesWUpdate.ps1**: Applies fixes for Windows Update.
- **Invoke-TeamsAddinFix.ps1**: Fixes Teams add-in issues.
- **Invoke-RebuildOST.ps1**: Rebuilds Outlook OST file.
- **Invoke-TeamsCacheClean.ps1**: Cleans Teams cache.
- **Invoke-TeamsRemoval.ps1**: Removes Teams.
- **Invoke-TeamsReset.ps1**: Resets Teams.

## XAML Files

### XAML/
- **mainWindow.xaml**: Defines the main window layout and controls.
- **FixEdgeWindow.xaml**: Defines the layout for the Edge fix window.
- **FixOutlookWindow.xaml**: Defines the layout for the Outlook fix window.
- **FixTeamsWindow.xaml**: Defines the layout for the Teams fix window.
- **FixWUpdateWindow.xaml**: Defines the layout for the Windows Update fix window.

## JSON Files

### JSON/
- **FixButtonMappings.json**: Maps buttons to their corresponding functions for the fix windows.
- **DNSList.json**: Contains a list of DNS providers for the DNS ComboBox.

## Installation

1. Clone the repository to your local machine.
2. Ensure you have PowerShell installed.
3. Open PowerShell and navigate to the project directory.

## Usage

1. Run the main script:
   ```powershell
   .\ESSToolBox.ps1
