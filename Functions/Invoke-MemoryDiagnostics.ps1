<#
.SYNOPSIS
Runs the Windows Memory Diagnostics tool.

.DESCRIPTION
The Invoke-MemoryDiagnostics function launches the Windows Memory Diagnostics tool (mdsched.exe) to check for memory problems. 
It provides feedback to the user by displaying a message before starting the diagnostics.

.PARAMETER None
This function does not take any parameters.

.NOTES
File Name      : Invoke-MemoryDiagnostics.ps1
The function uses the Start-Process cmdlet to launch the Windows Memory Diagnostics tool.
The function is part of the ToolBox project and is stored in the GitHub repository https://github.com/finkuja/ToolBox

.EXAMPLE
Invoke-MemoryDiagnostics
Prompts the user and runs the Windows Memory Diagnostics tool.

#>
function Invoke-MemoryDiagnostics {
    # Run Windows Memory Test
    Write-Host "Running Windows Memory Test..." -ForegroundColor Yellow
    Start-Process mdsched.exe
}