function Reset-EdgeProfile {
    # Create a confirmation dialog
    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to reset the Microsoft Edge browser profile?",
        "Confirmation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirmation -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host "Operation cancelled by the user."
        return
    }
    try {
        # Step 1: Reset Edge profile silently
        Write-Host "Resetting Edge browser profile..."
        Start-Process "msedge" -ArgumentList "--reset-profile" -NoNewWindow -Wait
        Write-Host "Edge browser profile has been reset." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to reset Edge profile. Error: $_" -ForegroundColor Red
    }
}