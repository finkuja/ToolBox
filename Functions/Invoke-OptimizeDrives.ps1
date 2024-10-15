<#
.SYNOPSIS
Optimizes the drives on the system by performing defragmentation on HDDs and retrim on SSDs.

.DESCRIPTION
The Invoke-OptimizeDrives function is designed to enhance the performance of the drives on the system by performing specific optimization tasks based on the type of drive. 
It prompts the user for confirmation before proceeding to ensure that the user is aware of the potential time-consuming nature of the operation.

The function iterates through all volumes on the system, identifies the type of each drive (SSD or HDD), and performs the appropriate optimization operation. 
For SSDs, it performs a retrim operation, which helps maintain the performance of the SSD by informing the drive of unused blocks that can be erased. 
For HDDs, it performs a defragmentation operation, which reorganizes fragmented data to improve read and write speeds.

The function provides verbose output to keep the user informed of the progress and status of the optimization process. 
It also includes error handling to manage any issues that arise during the optimization, providing feedback to the user through console messages. 
This ensures that the user is aware of any drives that could not be optimized and the reasons why.

Additionally, the function requires administrative privileges to run, as optimizing drives is a system-level operation. 
It uses Windows Forms to display a confirmation dialog box, making it user-friendly and interactive.

.PARAMETERS
None

.EXAMPLE
Invoke-OptimizeDrives

This example runs the Invoke-OptimizeDrives function, prompting the user for confirmation and then optimizing the drives
based on their media type.

.NOTES
- This function requires administrative privileges to run.
- The function uses Windows Forms to display a confirmation dialog box.
- The function handles errors and provides feedback to the user through console messages.

#>
function Invoke-OptimizeDrives {
    # Optimize Drives
    function Invoke-DeleteTempFiles {
        $confirmation = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to optimize the drives? This operation can take a long time depending on the type of drive.",
            "Confirmation",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    
        if ($confirmation -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-Host "Operation cancelled by the user."
            return
        }
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
} 