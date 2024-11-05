<#
.SYNOPSIS
    Installs or uninstalls packages using WinGet based on a JSON configuration file.

.DESCRIPTION
    The Invoke-WinGet function allows you to install or uninstall packages using WinGet. 
    The packages and their corresponding WinGet IDs are specified in a JSON file.

.PARAMETER Action
    Specifies the action to perform. Valid values are "Install" or "Uninstall".

.PARAMETER window
    A reference to the current window (used for UI interactions).

.PARAMETER PackageName
    The name of the package to install or uninstall. This should match a key in the JSON file.

.PARAMETER JsonFilePath
    The path to the JSON file containing package information. The JSON file should map package names to WinGet IDs.

.EXAMPLE
    # Path to the JSON file
    $jsonFilePath = "path\to\packages.json"

    # Example usage of Invoke-WinGet to install Google Chrome
    Invoke-WinGet -Action "Install" -window $window -PackageName "Google Chrome" -JsonFilePath $jsonFilePath

    # Example usage of Invoke-WinGet to uninstall Google Chrome
    Invoke-WinGet -Action "Uninstall" -window $window -PackageName "Google Chrome" -JsonFilePath $jsonFilePath

.NOTES
    The function checks if the JSON file exists at the specified path. If the file exists, it reads the content 
    and converts it from JSON format to a PowerShell object. The resulting hashtable is used to map package names 
    to their corresponding WinGet IDs. The function then uses the WinGet command-line tool to install or uninstall 
    the specified package(s).
#>
function Invoke-WinGet {
    param (
        [ValidateSet("Install", "Uninstall")]
        [string]$Action,
        [System.Windows.Window]$window,
        [string]$PackageName,
        [string]$JsonFilePath
    )

    # Load the packages from the JSON file
    if (Test-Path -Path $JsonFilePath) {
        $packages = Get-Content -Path $JsonFilePath | ConvertFrom-Json
    }
    else {
        [System.Windows.MessageBox]::Show("JSON file not found at path: $JsonFilePath")
        return
    }

    # Check if the package exists in the JSON file
    $packageExists = $false
    foreach ($key in $packages.PSObject.Properties.Name) {
        if ($key -eq $PackageName) {
            $packageExists = $true
            break
        }
    }

    if (-not $packageExists) {
        [System.Windows.MessageBox]::Show("Package '$PackageName' not found.")
        return
    }

    $packageId = $packages.$PackageName

    switch ($Action) {
        "Install" {
            # Install the specified package
            if ($packageId -is [array]) {
                foreach ($id in $packageId) {
                    Start-Process "winget" -ArgumentList "install --id $id -e --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait
                }
            }
            else {
                Start-Process "winget" -ArgumentList "install --id $packageId -e --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait
            }
        }
        "Uninstall" {
            # Uninstall the specified package
            if ($packageId -is [array]) {
                foreach ($id in $packageId) {
                    Start-Process "winget" -ArgumentList "uninstall --id $id -e" -NoNewWindow -Wait
                }
            }
            else {
                Start-Process "winget" -ArgumentList "uninstall --id $packageId -e" -NoNewWindow -Wait
            }
        }
    }
}