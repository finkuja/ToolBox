
# ESS ToolBox

ESS ToolBox is a PowerShell-based tool that provides various functionalities for system tweaks, installations, and fixes. The tool uses a graphical user interface (GUI) built on WPF with XAML and dynamically loads content based on user interactions. Please note that ESS ToolBox is intended to serve as a foundational base upon which you can build your own solutions. It is not officially supported, and users are expected to rely on their own expertise and resources when utilizing this tool.

## Folder Structure

```markdown
ESS ToolBox/
├── 

ESSToolBox.ps1


├── Functions/
│   ├── Show-ChildWindow.ps1
│   ├── Invoke-WinGet.ps1
│   ├── Invoke-Tweak.ps1
│   ├── Invoke-DeleteTempFiles.ps1
│   ├── Invoke-OptimizeDrives.ps1
│   ├── Invoke-DiskCleanup.ps1
│   ├── Reset-EdgeCache.ps1
│   ├── Reset-EdgeProfile.ps1
│   ├── Remove-Edge.ps1
│   ├── Invoke-FixesWUpdate.ps1
│   ├── Invoke-TeamsAddinFix.ps1
│   ├── Invoke-RebuildOST.ps1
│   ├── Invoke-TeamsCacheClean.ps1
│   ├── Invoke-TeamsRemoval.ps1
│   ├── Invoke-TeamsReset.ps1
├── XAML/
│   ├── mainWindow.xaml
│   ├── FixEdgeWindow.xaml
│   ├── FixOutlookWindow.xaml
│   ├── FixTeamsWindow.xaml
│   ├── FixWUpdateWindow.xaml
├── JSON/
│   ├── FixButtonMappings.json
│   ├── DNSList.json
```

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

## Prerequisites

1. Ensure PowerShell installed in the machine.
2. Run PowerShell with Administrator privileges.
2. The script can be executed in two modes:
    - **Online Mode**: This mode uses `Invoke-RestMethod` and `Invoke-Expression` to run the script directly from the web.
      ```powershell
      irm <<url to github raw file or your own url>> | iex
      ```
    - **Offline Mode**: This mode runs the script from your local machine.
      ```powershell
      .\ESSToolBox.ps1 -OfflineMode
      ```
3. Use the GUI to navigate through the tabs and perform various actions such as installing packages, applying tweaks, and fixing issues.

## Contributors

- [finkuja](https://github.com/finkuja)

## License

This project is licensed under the MIT License.
```
Copyright (c) 2024 Carlos Alvarez Magariños

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```
