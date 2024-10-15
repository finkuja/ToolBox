function Invoke-MemoryDiagnostics {
    # Run Windows Memory Test
    Write-Host "Running Windows Memory Test..." -ForegroundColor Yellow
    Start-Process mdsched.exe
}