function Invoke-DeleteTempFiles {
    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to delete all temporary files?",
        "Confirmation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirmation -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host "Operation cancelled by the user."
        return
    }
    # Delete temporary files
    Write-Host "Deleting temporary files..." -ForegroundColor Green

    $lockedFiles = @()
    $nonExistentPaths = @()

    function Remove-Files {
        param (
            [string[]]$paths,
            [string]$protectedFolder
        )

        foreach ($path in $paths) {
            if (Test-Path $path) {
                Get-ChildItem -Path $path -Recurse -Force | ForEach-Object {
                    # Skip the protected folder
                    if ($_.FullName -like "$protectedFolder*") {
                        return
                    }

                    try {
                        Remove-Item -Path $_.FullName -Force -Recurse -ErrorAction Stop
                    }
                    catch {
                        if ($_.Exception.Message -match "because it is being used by another process") {
                            $lockedFiles += $_.FullName
                        }
                        else {
                            $nonExistentPaths += $_.FullName
                        }
                    }
                }
            }
            else {
                $nonExistentPaths += $path
            }
        }
    }

    # Paths to remove files from
    $pathsToClean = @("C:\Windows\Temp", $env:TEMP)

    # Path to protect (ESSToolBox folder in the temp directory)
    $protectedFolder = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "ESSToolBox")

    # Remove files from specified paths, excluding the protected folder
    Remove-Files -paths $pathsToClean -protectedFolder $protectedFolder

    if ($lockedFiles.Count -gt 0) {
        Write-Host "The following files were not deleted because they are in use:" -ForegroundColor Yellow
        $lockedFiles | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
    }

    if ($nonExistentPaths.Count -gt 0) {
        Write-Host "The following paths do not exist:" -ForegroundColor Yellow
        $nonExistentPaths | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
    }

    if ($lockedFiles.Count -eq 0 -and $nonExistentPaths.Count -eq 0) {
        Write-Host "All temporary files were successfully deleted." -ForegroundColor Green
    }
}