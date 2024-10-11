function Invoke-Install {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.Windows.Window]$MainWindow
    )

    # Find the existing InstallTabGrid in the main window
    $installTabGrid = $MainWindow.FindName("InstallTabGrid")
    if ($null -eq $installTabGrid) {
        Write-Host "InstallTabGrid not found in the main window." -ForegroundColor Red
        return
    }

    $xaml = @"
        <Grid xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'>
            <Grid.RowDefinitions>
                <RowDefinition Height='Auto'/>
                <RowDefinition Height='Auto'/>
                <RowDefinition Height='Auto'/>
                <RowDefinition Height='Auto'/>
                <RowDefinition Height='Auto'/>
                <RowDefinition Height='Auto'/>
                <RowDefinition Height='Auto'/>
                <RowDefinition Height='Auto'/>
                <RowDefinition Height='Auto'/>
                <RowDefinition Height='*'/> <!-- This row will take the remaining space -->
                <RowDefinition Height='Auto'/> <!-- This row is for the buttons -->
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width='Auto'/>
                <ColumnDefinition Width='Auto'/>
                <ColumnDefinition Width='Auto'/>
                <ColumnDefinition Width='Auto'/>
            </Grid.ColumnDefinitions>
            <CheckBox Content='Adobe Creative Cloud' x:Name='AdobeCreativeCloud' Grid.Row='0' Grid.Column='0' Margin='10'/>
            <CheckBox Content='Adobe Reader DC 64-Bit' x:Name='AdobeAcrobat' Grid.Row='0' Grid.Column='1' Margin='10'/>
            <CheckBox Content='Google Chrome' x:Name='GoogleChrome' Grid.Row='0' Grid.Column='2' Margin='10'/>
            <CheckBox Content='Fiddler' x:Name='FiddlerClassic' Grid.Row='1' Grid.Column='0' Margin='10'/>
            <CheckBox Content='HWMonitor' x:Name='HWMonitor' Grid.Row='1' Grid.Column='1' Margin='10'/>
            <CheckBox Content='.NET All Versions' x:Name='MicrosoftDotNetRuntime' Grid.Row='1' Grid.Column='2' Margin='10'/>
            <CheckBox Content='Microsoft Edge' x:Name='MicrosoftEdge' Grid.Row='2' Grid.Column='0' Margin='10'/>
            <CheckBox Content='Microsoft Office 365' x:Name='Microsoft365AppsForEnterprise' Grid.Row='2' Grid.Column='1' Margin='10'/>
            <CheckBox Content='Microsoft OneDrive' x:Name='MicrosoftOneDrive' Grid.Row='2' Grid.Column='2' Margin='10'/>
            <CheckBox Content='Microsoft OneNote (UWP)' x:Name='MicrosoftOneNote' Grid.Row='3' Grid.Column='0' Margin='10'/>
            <CheckBox Content='Microsoft Teams' x:Name='MicrosoftTeams' Grid.Row='3' Grid.Column='1' Margin='10'/>
            <CheckBox Content='Mozilla Firefox' x:Name='MozillaFirefox' Grid.Row='3' Grid.Column='2' Margin='10'/>
            <CheckBox Content='Power Automate' x:Name='PowerAutomate' Grid.Row='4' Grid.Column='0' Margin='10'/>
            <CheckBox Content='Power BI Desktop' x:Name='PowerBIDesktop' Grid.Row='4' Grid.Column='1' Margin='10'/>
            <CheckBox Content='PowerToys' x:Name='PowerToys' Grid.Row='4' Grid.Column='2' Margin='10'/>
            <CheckBox Content='Quick Assist' x:Name='QuickAssist' Grid.Row='5' Grid.Column='0' Margin='10'/>
            <CheckBox Content='Remote Desktop' x:Name='MicrosoftRemoteDesktop' Grid.Row='5' Grid.Column='1' Margin='10'/>
            <CheckBox Content='SARA Tool' x:Name='MicrosoftSupportAndRecoveryAssistant' Grid.Row='5' Grid.Column='2' Margin='10'/>
            <CheckBox Content='Surface Diagnostic Toolkit' x:Name='SurfaceDiagnosticToolkit' Grid.Row='6' Grid.Column='0' Margin='10'/>
            <CheckBox Content='Visio Viewer 2016' x:Name='MicrosoftVisioViewer' Grid.Row='6' Grid.Column='1' Margin='10'/>
            <CheckBox Content='Visual Studio Code' x:Name='MicrosoftVisualStudioCode' Grid.Row='6' Grid.Column='2' Margin='10'/>
            <CheckBox Content='7-Zip' x:Name='SevenZip' Grid.Row='7' Grid.Column='0' Margin='10'/>
            <CheckBox Content='Show All Installed' x:Name='ShowInstalled' Grid.Row='7' Grid.Column='1' Margin='10'/>
            <Button Content='Install' x:Name='InstallButton' Grid.Row='10' Grid.Column='0' Margin='10'/>
            <Button Content='Uninstall' x:Name='UninstallButton' Grid.Row='10' Grid.Column='1' Margin='10'/>
            <Button Content='Get Installed' x:Name='GetPackagesButton' Grid.Row='10' Grid.Column='2' Margin='10'/>
            <Button Content='Check All' x:Name='CheckAllButton' Grid.Row='10' Grid.Column='3' Margin='10'/>
        </Grid>
"@

    #################
    ## End of XAML ##
    #################

    # Load the XAML
    try {
        $reader = (New-Object System.Xml.XmlTextReader -ArgumentList (New-Object System.IO.StringReader -ArgumentList $xaml))
        $installTabContent = [Windows.Markup.XamlReader]::Load($reader)
    }
    catch {
        Write-Host "Failed to load XAML: $_" -ForegroundColor Red
        return
    }

    if ($null -eq $installTabContent) {
        Write-Host "Failed to load InstallTab content." -ForegroundColor Red
        return
    }

    # Add the loaded content to the InstallTabGrid
    $installTabGrid.Children.Clear()
    $installTabGrid.Children.Add($installTabContent)

    # Find controls
    $installButton = $installTabContent.FindName("InstallButton")
    $uninstallButton = $installTabContent.FindName("UninstallButton")
    $getPackagesButton = $installTabContent.FindName("GetPackagesButton")
    $checkAllButton = $installTabContent.FindName("CheckAllButton")

    # Ensure buttons are found
    if ($null -eq $installButton -or $null -eq $uninstallButton -or $null -eq $getPackagesButton -or $null -eq $checkAllButton) {
        Write-Host "One or more buttons not found in the XAML." -ForegroundColor Red
        return
    }

    # Recursive function to find all checkboxes
    function Get-AllCheckboxes {
        param (
            [System.Windows.DependencyObject]$parent
        )
        $checkboxes = @()
        for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($parent); $i++) {
            $child = [System.Windows.Media.VisualTreeHelper]::GetChild($parent, $i)
            if ($child -is [System.Windows.Controls.CheckBox]) {
                $checkboxes += $child
            }
            $checkboxes += Get-AllCheckboxes -parent $child
        }
        return $checkboxes
    }

    # Find all checkboxes
    $checkboxes = Get-AllCheckboxes -parent $installTabContent

    # Set the initial state of all checkboxes to checked
    foreach ($checkbox in $checkboxes) {
        $checkbox.IsChecked = $true
    }

    function Invoke-WingetCommand {
        param (
            [string]$action,
            [string]$id
        )
        switch ($action) {
            "install" {
                Write-Host "Installing package: $id"
                Install-WinGetPackage -Id $id -AcceptSourceAgreements -AcceptPackageAgreements -Interactive
            }
            "uninstall" {
                Write-Host "Uninstalling package: $id"
                Uninstall-WinGetPackage -Id $id -AcceptSourceAgreements -AcceptPackageAgreements
            }
        }
    }

    $checkboxActions = @{
        "AdobeAcrobat"                         = "Adobe.Acrobat.Reader.64-bit"
        "AdobeCreativeCloud"                   = "Adobe.CreativeCloud"
        "GoogleChrome"                         = "Google.Chrome"
        "MicrosoftEdge"                        = "Microsoft.Edge"
        "FiddlerClassic"                       = "Telerik.Fiddler.Classic"
        "MozillaFirefox"                       = "Mozilla.Firefox"
        "HWMonitor"                            = "CPUID.HWMonitor"
        "MicrosoftSupportAndRecoveryAssistant" = "Microsoft.SupportAndRecoveryAssistant"
        "MicrosoftDotNetRuntime"               = @("Microsoft.DotNet.DesktopRuntime.3_1", "Microsoft.DotNet.DesktopRuntime.5", "Microsoft.DotNet.DesktopRuntime.6", "Microsoft.DotNet.DesktopRuntime.7", "Microsoft.DotNet.DesktopRuntime.8")
        "Microsoft365AppsForEnterprise"        = "Microsoft.Office"
        "MicrosoftOneDrive"                    = "Microsoft.OneDrive"
        "MicrosoftOneNote"                     = "XPFFZHVGQWWLHB"
        "PowerAutomate"                        = "9NFTCH6J7FHV"
        "PowerBIDesktop"                       = "Microsoft.PowerBI"
        "PowerToys"                            = "Microsoft.PowerToys"
        "QuickAssist"                          = "9P7BP5VNWKX5"
        "MicrosoftRemoteDesktop"               = "9WZDNCRFJ3PS"
        "SurfaceDiagnosticToolkit"             = "9NF1MR6C60ZF"
        "MicrosoftTeams"                       = "Microsoft.Teams"
        "MicrosoftVisioViewer"                 = "Microsoft.VisioViewer"
        "MicrosoftVisualStudioCode"            = "Microsoft.VisualStudioCode"
        "SevenZip"                             = "7zip.7zip"
    }

    # Define the action for the Install button
    $installButton.Add_Click({
            Write-Host "Install button clicked"
            $selectedCheckboxes = $checkboxes | Where-Object { $_.IsChecked -eq $true }
            foreach ($checkbox in $selectedCheckboxes) {
                $id = $checkboxActions[$checkbox.Name]
                if ($id -is [Array]) {
                    foreach ($subId in $id) {
                        Invoke-WingetCommand -action "install" -id $subId
                    }
                }
                else {
                    Invoke-WingetCommand -action "install" -id $id
                }
            }
        })

    # Define the action for the Uninstall button
    $uninstallButton.Add_Click({
            Write-Host "Uninstall button clicked"
            $selectedCheckboxes = $checkboxes | Where-Object { $_.IsChecked -eq $true }
            foreach ($checkbox in $selectedCheckboxes) {
                $id = $checkboxActions[$checkbox.Name]
                if ($id -is [Array]) {
                    foreach ($subId in $id) {
                        Invoke-WingetCommand -action "uninstall" -id $subId
                    }
                }
                else {
                    Invoke-WingetCommand -action "uninstall" -id $id
                }
            }
        })

    # Define the action for the Get Installed button
    $getPackagesButton.Add_Click({
            Write-Host "Get Installed button clicked"
            # Run the Get-WinGetPackage command and capture the output directly
            $output = Get-WinGetPackage

            # Extract the package names from the output
            $packageNames = $output | Select-Object -ExpandProperty Name

            # Iterate through the checkboxes and check the ones that match the installed software
            foreach ($checkbox in $checkboxes) {
                $checkboxName = $checkbox.Name
                if ($packageNames -contains $checkboxName) {
                    $checkbox.IsChecked = $true
                }
                else {
                    $checkbox.IsChecked = $false
                }
            }
        })

    # Define the action for the Check All button
    $checkAllButton.Add_Click({
            Write-Host "Check All button clicked"
            $allChecked = $true

            # Check the state of each checkbox
            foreach ($checkbox in $checkboxes) {
                if ($checkbox.IsChecked -eq $false) {
                    $allChecked = $false
                    break
                }
            }

            # Set the state of each checkbox
            foreach ($checkbox in $checkboxes) {
                $checkbox.IsChecked = -not $allChecked
            }

            # Update button content
            if ($allChecked) {
                $checkAllButton.Content = "Check All"
            }
            else {
                $checkAllButton.Content = "Uncheck All"
            }
        })

    Export-ModuleMember -Function Invoke-Install
}