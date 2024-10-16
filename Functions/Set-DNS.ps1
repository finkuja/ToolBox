<#
.SYNOPSIS
Sets or resets the DNS server addresses for all active network adapters.

.DESCRIPTION
The Set-DNS function configures the DNS server addresses for all active network adapters on the system. 
It can either set specific DNS addresses or reset the DNS settings to their default values.

.PARAMETER dnsAddresses
Specifies the DNS server addresses to set. This parameter is an array of strings.

.PARAMETER reset
A switch parameter that, when specified, resets the DNS settings to their default values.

.NOTES
File Name      : Set-DNS.ps1
The function uses the Get-NetAdapter and Set-DnsClientServerAddress cmdlets to manage DNS settings.
The function is part of the ToolBox project and is stored in the GitHub repository https://github.com/finkuja/ToolBox

.EXAMPLE
Set-DNS -dnsAddresses "8.8.8.8", "8.8.4.4"
Sets the DNS server addresses to 8.8.8.8 and 8.8.4.4 for all active network adapters.

.EXAMPLE
Set-DNS -reset
Resets the DNS settings to their default values for all active network adapters.

#>
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