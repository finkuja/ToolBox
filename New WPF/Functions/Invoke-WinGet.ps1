function Invoke-WinGet { 
    param (
        [ValidateSet("Install", "Uninstall")]
        [string]$Action,
        [System.Windows.Window]$window,
        [string]$PackageName
    )

    # Define the list of packages and their corresponding IDs
    $packages = @{
        "Adobe Creative Cloud"                     = "Adobe.CreativeCloud"
        "Adobe Acrobat"                            = "Adobe.Acrobat.Reader.64-bit"
        "Google Chrome"                            = "Google.Chrome"
        "Fiddler Classic"                          = "Telerik.Fiddler.Classic"
        "HWMonitor"                                = "CPUID.HWMonitor"
        "Microsoft .Net Runtime"                   = @(
            "Microsoft.DotNet.DesktopRuntime.3_1",
            "Microsoft.DotNet.DesktopRuntime.5",
            "Microsoft.DotNet.DesktopRuntime.6",
            "Microsoft.DotNet.DesktopRuntime.7",
            "Microsoft.DotNet.DesktopRuntime.8"
        )
        "Microsoft Edge"                           = "Microsoft.Edge"
        "Microsoft 365 Apps for Enterprise"        = "Microsoft.Office"
        "Microsoft OneDrive"                       = "Microsoft.OneDrive"
        "Microsoft OneNote"                        = "XPFFZHVGQWWLHB"
        "Microsoft Teams"                          = "Microsoft.Teams"
        "Mozilla Firefox"                          = "Mozilla.Firefox"
        "Power Automate"                           = "9NFTCH6J7FHV"
        "Power BI Desktop"                         = "Microsoft.PowerBI"
        "PowerToys"                                = "Microsoft.PowerToys"
        "Quick Assist"                             = "9P7BP5VNWKX5"
        "Microsoft Remote Desktop"                 = "9WZDNCRFJ3PS"
        "Microsoft Support and Recovery Assistant" = "Microsoft.SupportAndRecoveryAssistant"
        "Surface Diagnostic Toolkit"               = "9NF1MR6C60ZF"
        "Microsoft VisioViewer"                    = "Microsoft.VisioViewer"
        "Microsoft Visual Studio Code"             = "Microsoft.VisualStudioCode"
        "7-Zip"                                    = "7zip.7zip"
    }

    switch ($Action) {
        "Install" {
            # Install the specified package
            if ($packages.ContainsKey($PackageName)) {
                $packageId = $packages[$PackageName]
                if ($packageId -is [array]) {
                    foreach ($id in $packageId) {
                        Start-Process "winget" -ArgumentList "install --id $id -e --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait
                    }
                }
                else {
                    Start-Process "winget" -ArgumentList "install --id $packageId -e --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait
                }
                [System.Windows.MessageBox]::Show("Package '$PackageName' has been installed.")
            }
            else {
                [System.Windows.MessageBox]::Show("Package '$PackageName' not found.")
            }
        }
        "Uninstall" {
            # Uninstall the specified package
            if ($packages.ContainsKey($PackageName)) {
                $packageId = $packages[$PackageName]
                if ($packageId -is [array]) {
                    foreach ($id in $packageId) {
                        Start-Process "winget" -ArgumentList "uninstall --id $id -e" -NoNewWindow -Wait
                    }
                }
                else {
                    Start-Process "winget" -ArgumentList "uninstall --id $packageId -e" -NoNewWindow -Wait
                }
                [System.Windows.MessageBox]::Show("Package '$PackageName' has been uninstalled.")
            }
            else {
                [System.Windows.MessageBox]::Show("Package '$PackageName' not found.")
            }
        }
        
    }
}