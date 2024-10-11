Add-Type -AssemblyName PresentationFramework

# Define the XAML content
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="MainWindow" Height="350" Width="525">
    <StackPanel>
        <Button Name="CheckAllButton" Content="Check All" Width="100" Height="30" Margin="10"/>
        <CheckBox Name="CheckBox1" Content="Option 1" Margin="10"/>
        <CheckBox Name="CheckBox2" Content="Option 2" Margin="10"/>
        <CheckBox Name="CheckBox3" Content="Option 3" Margin="10"/>
    </StackPanel>
</Window>
"@

# Load the XAML content
$reader = (New-Object System.Xml.XmlNodeReader ([xml]$xaml))
$window = [Windows.Markup.XamlReader]::Load($reader)

# Find controls
$checkAllButton = $window.FindName("CheckAllButton")
$checkBox1 = $window.FindName("CheckBox1")
$checkBox2 = $window.FindName("CheckBox2")
$checkBox3 = $window.FindName("CheckBox3")

# Define the action for the Check All button
$checkAllButton.Add_Click({
        Write-Host "Check All button clicked"
        $allChecked = $true

        # Check the state of each checkbox
        foreach ($checkBox in @($checkBox1, $checkBox2, $checkBox3)) {
            if ($checkBox.IsChecked -eq $false) {
                $allChecked = $false
                break
            }
        }

        # Set the state of each checkbox
        foreach ($checkBox in @($checkBox1, $checkBox2, $checkBox3)) {
            $checkBox.IsChecked = -not $allChecked
        }

        # Update button content
        $checkAllButton.Content = if ($allChecked) { "Check All" } else { "Uncheck All" }
    })

# Show the window
$window.ShowDialog() | Out-Null