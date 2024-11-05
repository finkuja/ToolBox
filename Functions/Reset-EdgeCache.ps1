<#
.SYNOPSIS
Resets the Microsoft Edge cache.

.DESCRIPTION
The Reset-EdgeCache function clears the cache for Microsoft Edge by simulating key presses to navigate to the settings page for clearing browsing data. 
It opens the Edge browser, waits for it to load, and then uses the WScript.Shell COM object to send the necessary key presses.

.PARAMETER None
This function does not take any parameters.

.NOTES
File Name      : Reset-EdgeCache.ps1
The function uses the Start-Process cmdlet to open Edge and the WScript.Shell COM object to send key presses.
The function is part of the ToolBox project and is stored in the GitHub repository https://github.com/finkuja/ToolBox

.EXAMPLE
Reset-EdgeCache
Opens Microsoft Edge and navigates to the settings page to clear the browsing data.

#>
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
