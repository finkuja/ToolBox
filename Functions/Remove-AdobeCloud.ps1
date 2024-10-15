function Remove-AdobeCloud {
    # Create a confirmation dialog
    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to remove Adobe Cloud?",
        "Confirmation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirmation -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host "Operation cancelled by the user."
        return
    }

    # Remove Adobe Cloud
    Write-Host "Removing Adobe Cloud..."
    # Code snippet from https://github.com/ChrisTitusTech/winutil/blob/main/docs/dev/features/Fixes/RunAdobeCCCleanerTool.md
    [string]$url = "https://swupmf.adobe.com/webfeed/CleanerTool/win/AdobeCreativeCloudCleanerTool.exe"

    Write-Host "The Adobe Creative Cloud Cleaner tool is hosted at"
    Write-Host "$url"

    try {
        # Don't show the progress because it will slow down the download speed
        $ProgressPreference = 'SilentlyContinue'

        Invoke-WebRequest -Uri $url -OutFile "$env:TEMP\AdobeCreativeCloudCleanerTool.exe" -UseBasicParsing -ErrorAction SilentlyContinue -Verbose

        # Revert back the ProgressPreference variable to the default value since we got the file desired
        $ProgressPreference = 'Continue'

        Start-Process -FilePath "$env:TEMP\AdobeCreativeCloudCleanerTool.exe" -Wait -ErrorAction SilentlyContinue -Verbose
    }
    catch {
        Write-Error $_.Exception.Message
    }
    finally {
        if (Test-Path -Path "$env:TEMP\AdobeCreativeCloudCleanerTool.exe") {
            Write-Host "Cleaning up..."
            Remove-Item -Path "$env:TEMP\AdobeCreativeCloudCleanerTool.exe" -Verbose
        }
    }
}