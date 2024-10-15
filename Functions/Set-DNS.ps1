# Define the Set-DNS function
function Set-DNS {
    param (
        [string[]]$dnsAddresses,
        [switch]$reset
    )

    $networkAdapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }
    foreach ($adapter in $networkAdapters) {
        if ($reset) {
            Set-DnsClientServerAddress -InterfaceAlias $adapter.Name -ResetServerAddresses
            Write-Host "DNS settings reset to default for adapter: $($adapter.Name)"
        }
        else {
            Set-DnsClientServerAddress -InterfaceAlias $adapter.Name -ServerAddresses $dnsAddresses
            Write-Host "DNS settings applied for adapter: $($adapter.Name)"
        }
    }
}