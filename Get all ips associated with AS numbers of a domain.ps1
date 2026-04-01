# -----------------------------
# Script: DNS → IP → ASN → Prefixes
#
# Use .\"Find all ips associated with AS numbers of a domain.ps1 -Hostname google.com"
# -----------------------------

param (
    [Parameter(Mandatory = $true)]
    [string]$Hostname
)

try {
    # Step 1: Resolve DNS to IP
    $ip = [System.Net.Dns]::GetHostAddresses($Hostname) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1

    if (-not $ip) {
        Write-Error "No IPv4 address found for $Hostname"
        exit
    }

    Write-Host "Resolved $Hostname → $ip"

    # Step 2: Get ASN using RIPE Stat API
    $asnData = Invoke-RestMethod "https://stat.ripe.net/data/prefix-overview/data.json?resource=$ip"
    $asn = $asnData.data.asns.asn
    $holder = $asnData.data.asns.holder

    Write-Host "IP $ip belongs to ASN $asn ($holder)"

    # Step 3: Get all prefixes announced by ASN via RIPE Stat API
    $prefixData = (Invoke-RestMethod "https://stat.ripe.net/data/announced-prefixes/data.json?resource=AS$asn").data.prefixes

    if (-not $prefixData) {
        Write-Error "No prefixes found for ASN $asn"
        exit
    }

    # Step 4: Output the prefixes
    $prefixData | ForEach-Object {
        [PSCustomObject]@{
            ASN    = "AS$asn"
            Prefix = $_.prefix
        }
    } | Format-Table -AutoSize

} catch {
    Write-Error "An error occurred: $_"
}
