function Invoke-DeleteTempFiles {
    # Delete temporary files
    Write-Host "Deleting temporary files..." -ForegroundColor Green

    $lockedFiles = @()
    $nonExistentPaths = @()

    function Remove-Files {
        param (
            [string[]]$paths
        )

        foreach ($path in $paths) {
            if (Test-Path $path) {
                Get-ChildItem -Path $path -Recurse -Force | ForEach-Object {
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

    # Remove files from specified paths
    Remove-Files -paths $pathsToClean

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