function Invoke-Install {
    [CmdletBinding()]
    param (
        [switch]$DisableInstall,
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
            <Button Content='Install' x:Name='InstallButton' Grid.Row='10' Grid.Column='0' Margin='10' Background='#0D9D3B' Foreground='White' HorizontalAlignment='Stretch' VerticalAlignment='Stretch' Padding='30,15' Cursor='Hand'>
                <Button.Template>
                    <ControlTemplate TargetType='Button'>
                        <Border Background='{TemplateBinding Background}' CornerRadius='5'>
                            <ContentPresenter HorizontalAlignment='Stretch' VerticalAlignment='Stretch'/>
                        </Border>
                    </ControlTemplate>
                </Button.Template>
            </Button>
            <Button Content='Uninstall' x:Name='UninstallButton' Grid.Row='10' Grid.Column='1' Margin='10' Background='#E30101' Foreground='White' HorizontalAlignment='Stretch' VerticalAlignment='Stretch' Padding='30,15' Cursor='Hand'>
                <Button.Template>
                    <ControlTemplate TargetType='Button'>
                        <Border Background='{TemplateBinding Background}' CornerRadius='5'>
                            <ContentPresenter HorizontalAlignment='Stretch' VerticalAlignment='Stretch'/>
                        </Border>
                    </ControlTemplate>
                </Button.Template>
            </Button>
            <Button Content='Get Installed' x:Name='GetPackagesButton' Grid.Row='10' Grid.Column='2' Margin='10' Background='#185FF8' Foreground='White' HorizontalAlignment='Stretch' VerticalAlignment='Stretch' Padding='30,15' Cursor='Hand'>
                <Button.Template>
                    <ControlTemplate TargetType='Button'>
                        <Border Background='{TemplateBinding Background}' CornerRadius='5'>
                            <ContentPresenter HorizontalAlignment='Stretch' VerticalAlignment='Stretch'/>
                        </Border>
                    </ControlTemplate>
                </Button.Template>
            </Button>
            <Button Content='Check All' x:Name='CheckAllButton' Grid.Row='10' Grid.Column='3' Margin='10' Background='#185FF8' Foreground='White' HorizontalAlignment='Stretch' VerticalAlignment='Stretch' Padding='30,15' Cursor='Hand'>
                <Button.Template>
                    <ControlTemplate TargetType='Button'>
                        <Border Background='{TemplateBinding Background}' CornerRadius='5'>
                            <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'/>
                        </Border>
                    </ControlTemplate>
                </Button.Template>
            </Button>
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

    # Disable the Install Tab if $DisableInstall is $true
    if ($DisableInstall) {
        foreach ($control in $installTabContent.Children) {
            if ($control -is [System.Windows.Controls.CheckBox]) {
                $control.IsEnabled = $false
            }
        }
        $installButton.IsEnabled = $false
        $uninstallButton.IsEnabled = $false
        $getPackagesButton.IsEnabled = $false
        $checkAllButton.IsEnabled = $false
    }

    function Invoke-WingetCommand($action, $id) {
        Start-Process "winget" -ArgumentList "$action --id $id -e --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait
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
            $checkboxes = $installTabContent.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_.IsChecked }
            foreach ($checkbox in $checkboxes) {
                $id = $checkboxActions[$checkbox.Name]
                if ($id -is [Array]) {
                    foreach ($subId in $id) {
                        Invoke-WingetCommand "install" $subId
                    }
                }
                else {
                    Invoke-WingetCommand "install" $id
                }
            }
        })

    # Define the action for the Uninstall button
    $uninstallButton.Add_Click({
            $checkboxes = $installTabContent.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_.IsChecked }
            foreach ($checkbox in $checkboxes) {
                $id = $checkboxActions[$checkbox.Name]
                if ($id -is [Array]) {
                    foreach ($subId in $id) {
                        Invoke-WingetCommand "uninstall" $subId
                    }
                }
                else {
                    Invoke-WingetCommand "uninstall" $id
                }
            }
        })

    # Define the action for the Get Installed button
    $getPackagesButton.Add_Click({
            Start-Process "winget" -ArgumentList "list" -NoNewWindow -Wait
        })

    # Define the action for the Check All button
    $checkAllButton.Add_Click({
            foreach ($control in $installTabContent.Children) {
                if ($control -is [System.Windows.Controls.CheckBox]) {
                    $control.IsChecked = $true
                }
            }
        })
}
Export-ModuleMember -Function Invoke-Install