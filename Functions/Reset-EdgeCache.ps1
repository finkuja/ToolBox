function Reset-EdgeCache {
    try {
        # Open Edge browser
        $edgeProcess = Start-Process "msedge" -ArgumentList "about:blank" -PassThru

        # Wait for Edge to open
        Start-Sleep -Seconds 3

        # Simulate key presses to navigate to the settings page for clearing browsing data
        $shell = New-Object -ComObject "WScript.Shell"
        $shell.AppActivate($edgeProcess.Id)
        Start-Sleep -Milliseconds 500
        $shell.SendKeys("^+{DEL}")  # Ctrl+Shift+Delete to open the Clear browsing data dialog
        Start-Sleep -Milliseconds 500
        $shell.SendKeys("{ENTER}")  # Press Enter to confirm

        Write-Host "Edge settings page for clearing browsing data has been opened." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to open Edge settings page. Error: $_" -ForegroundColor Red
    }
}
