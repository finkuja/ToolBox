<#
.SYNOPSIS
Removes Microsoft Edge and all cached files from the system.

.DESCRIPTION
The Remove-Edge function performs the following actions:
1. Prompts the user with a message box to confirm the removal of Microsoft Edge and all cached files.
2. If confirmed, checks if Microsoft Edge is running and closes it.
3. Removes Edge cache and temporary files.
4. Uninstalls Microsoft Edge using the provided Uninstall-EdgeBrowser function.
5. Displays a message box with any files/folders that could not be removed.
6. Displays a summary of any errors that occurred during the process.

.PARAMETERS
None.

.EXAMPLE
PS> Remove-Edge

.NOTES
- This function requires the Uninstall-EdgeBrowser function to be defined elsewhere in the script or module.
- The function uses Windows Forms for message boxes, which requires the appropriate assemblies to be loaded.

#> 
function Remove-Edge {
    $errors = @()
    try {
        # Show a message box to advise the user
        $result = [System.Windows.Forms.MessageBox]::Show(
            "This action will fully remove Microsoft Edge and all cached files from the system. Do you want to proceed?",
            "Warning",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            # Step 1: Check if Edge is running and close it
            $edgeProcesses = Get-Process -Name "msedge" -ErrorAction SilentlyContinue
            if ($edgeProcesses) {
                Write-Host "Microsoft Edge is running. Closing it..."
                $edgeProcesses | ForEach-Object { $_.Kill() }
                Write-Host "Microsoft Edge has been closed." -ForegroundColor Green
            }

            # Step 2: Remove Edge cache and temporary files
            Write-Host "Removing Edge cache and temporary files..."
            $edgeCachePaths = @(
                "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache",
                "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Code Cache",
                "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\GPUCache",
                "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Media Cache",
                "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\ShaderCache"
            )
            $failedRemovals = @()
            foreach ($path in $edgeCachePaths) {
                try {
                    if (Test-Path $path) {
                        Remove-Item -Path $path -Recurse -Force
                        Write-Host "Removed: $path" -ForegroundColor Green
                    }
                    else {
                        Write-Host "Path not found: $path" -ForegroundColor Yellow
                    }
                }
                catch {
                    Write-Host "Failed to remove: $path" -ForegroundColor Red
                    $failedRemovals += $path
                }
            }

            # Step 3: Uninstall Edge using the provided function
            Write-Host "Uninstalling Edge browser..."
            try {
                Uninstall-EdgeBrowser
                Write-Host "Edge browser has been uninstalled." -ForegroundColor Green
            }
            catch {
                $errorMessage = "Failed to uninstall Edge browser. Error: $_"
                Write-Host $errorMessage -ForegroundColor Red
                $errors += $errorMessage
            }

            # Prompt user with the list of files/folders that could not be removed
            if ($failedRemovals.Count -gt 0) {
                $failedRemovalsMessage = "The following files/folders could not be removed:`n" + ($failedRemovals -join "`n")
                [System.Windows.Forms.MessageBox]::Show($failedRemovalsMessage, "Removal Errors", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        }
        else {
            Write-Host "Action canceled by the user." -ForegroundColor Yellow
        }
    }
    catch {
        $errorMessage = "Failed to reset Edge browser. Error: $_"
        Write-Host $errorMessage -ForegroundColor Red
        $errors += $errorMessage
    }
    finally {
        if ($errors.Count -gt 0) {
            $errorSummary = "The following errors occurred during the reset process:`n" + ($errors -join "`n")
            [System.Windows.Forms.MessageBox]::Show($errorSummary, "Errors", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}