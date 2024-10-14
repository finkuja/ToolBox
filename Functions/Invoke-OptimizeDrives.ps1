function Invoke-OptimizeDrives {
    # Optimize Drives
    Write-Host "Optimizing drives..." -ForegroundColor Green
    Get-Volume | Where-Object { $_.DriveLetter } | ForEach-Object { 
        try {
            $partition = Get-Partition -DriveLetter $_.DriveLetter -ErrorAction SilentlyContinue
            if ($null -eq $partition) {
                Write-Host "No partition found for drive $($_.DriveLetter). Skipping optimization." -ForegroundColor Yellow
                return
            }

            $physicalDisk = Get-PhysicalDisk | Where-Object { $_.DeviceID -eq $partition.DiskNumber }
            if ($null -eq $physicalDisk) {
                Write-Host "No physical disk found for partition $($partition.DiskNumber). Skipping optimization." -ForegroundColor Yellow
                return
            }

            $mediaType = $physicalDisk.MediaType
            if ($mediaType -eq 'SSD' -or $mediaType -eq 'Solid State Drive') {
                Optimize-Volume -DriveLetter $_.DriveLetter -ReTrim -Verbose
                Write-Host "SSD Drive $($_.DriveLetter) has been optimized with ReTrim." -ForegroundColor Green
            }
            elseif ($mediaType -eq 'HDD' -or $mediaType -eq 'Hard Disk Drive') {
                Optimize-Volume -DriveLetter $_.DriveLetter -Defrag -Verbose
                Write-Host "HDD Drive $($_.DriveLetter) has been optimized with Defrag." -ForegroundColor Green
            }
            else {
                Write-Host "Unknown media type for drive $($_.DriveLetter). Skipping optimization." -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host "Failed to optimize drive $($_.DriveLetter). Error: $_" -ForegroundColor Yellow
        }
    }
    Write-Host "Drive optimization process completed." -ForegroundColor Green
}