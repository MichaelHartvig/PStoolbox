function Get-ASN {
    param([string]$ip)
    (Invoke-RestMethod "https://stat.ripe.net/data/prefix-overview/data.json?resource=$ip").data.asns
}

Get-ASN 8.8.8.8
