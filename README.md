# ESS Tool

`EssTool.ps1` is a PowerShell script designed to provide a graphical user interface (GUI) for installing various Microsoft software packages using the `winget` package manager. The script creates a form with checkboxes for different software packages and an install button to initiate the installation process.

## Features

- Graphical User Interface (GUI) for selecting software packages.
- Checkboxes for Microsoft Office 365, PowerToys, Microsoft Teams, and Microsoft OneNote.
- Install button to start the installation of selected packages.

## Prerequisites

- Windows operating system.
- PowerShell 5.1 or later.
- `winget` package manager installed.

## Installation

1. Ensure that you have `winget` installed on your system. If not, you can install it from the [Microsoft Store](https://aka.ms/getwinget).
2. Download the `EssTool.ps1` script to your local machine.

## Usage

1. Open PowerShell with administrative privileges.
2. Navigate to the directory where `EssTool.ps1` is located.
3. Run the script using the following command:
   ```powershell
   .\EssTool.ps1
   ```
4. The GUI will appear with checkboxes for the following software packages:
   - Microsoft Office 365
   - PowerToys
   - Microsoft Teams
   - Microsoft OneNote
5. Select the software packages you want to install by checking the corresponding checkboxes.
6. Click the "Install" button to start the installation process for the selected packages.

## Script Details

The script creates checkboxes for each software package and an install button. When the install button is clicked, the script checks which checkboxes are selected and runs the `winget` command to install the corresponding packages.

### Checkboxes

- **Microsoft Office 365**
  - Checkbox Name: `Microsoft.Office`
  - Winget ID: `Microsoft.Office`

- **PowerToys**
  - Checkbox Name: `Microsoft.PowerToys`
  - Winget ID: `Microsoft.PowerToys`

- **Microsoft Teams**
  - Checkbox Name: `Microsoft.Teams`
  - Winget ID: `Microsoft.Teams`

- **Microsoft OneNote**
  - Checkbox Name: `Microsoft.OneNote`
  - Winget ID: `Microsoft.OneNote`

### Install Button

The install button is configured to run the `winget` command for each selected package. The command is executed in a new process without opening a new window, and the script waits for the installation to complete before proceeding.

## Contributing

If you would like to contribute to this project, please fork the repository and submit a pull request with your changes.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
```

This `README.md` file provides an overview of the script, its features, prerequisites, installation instructions, usage details, and information about contributing and licensing. Adjust the content as needed to fit your specific requirements.
