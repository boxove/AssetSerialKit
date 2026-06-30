# Script author: NaF

param(
    [int]$TimeoutMilliseconds = 400,
    [int]$MaxHostsPerSubnet = 254
)

$ErrorActionPreference = 'Continue'

$lanListPath = Join-Path $PSScriptRoot 'computers_lan.txt'
$logPath = Join-Path $PSScriptRoot 'lan_collect.log'

function Convert-IPv4ToUInt32([string]$address) {
    $bytes = [System.Net.IPAddress]::Parse($address).GetAddressBytes()
    [Array]::Reverse($bytes)
    return [BitConverter]::ToUInt32($bytes, 0)
}

function Convert-UInt32ToIPv4([uint32]$value) {
    $bytes = [BitConverter]::GetBytes($value)
    [Array]::Reverse($bytes)
    $address = New-Object System.Net.IPAddress -ArgumentList (,$bytes)
    return $address.ToString()
}

function Test-HostOnline([string]$address, [int]$timeout) {
    try {
        $ping = New-Object System.Net.NetworkInformation.Ping
        $reply = $ping.Send($address, $timeout)
        return $reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success
    } catch {
        return $false
    }
}

'Discovering LAN hosts...' | Tee-Object -FilePath $logPath

$targets = New-Object System.Collections.Generic.HashSet[string]
$adapters = @(Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true })

foreach ($adapter in $adapters) {
    for ($i = 0; $i -lt $adapter.IPAddress.Count; $i++) {
        $ip = $adapter.IPAddress[$i]
        $mask = $adapter.IPSubnet[$i]
        if (-not $ip -or $ip -match ':' -or -not $mask -or $mask -match ':') { continue }

        try {
            $ipValue = Convert-IPv4ToUInt32 $ip
            $maskValue = Convert-IPv4ToUInt32 $mask
            $network = $ipValue -band $maskValue
            $broadcast = $network -bor (-bnot $maskValue)
            $first = $network + 1
            $last = $broadcast - 1
            $count = [math]::Min(($last - $first + 1), $MaxHostsPerSubnet)

            ('Scanning subnet ' + $ip + '/' + $mask + ' hosts: ' + $count) | Tee-Object -FilePath $logPath -Append
            for ($offset = 0; $offset -lt $count; $offset++) {
                $target = Convert-UInt32ToIPv4 ([uint32]($first + $offset))
                if ($target -eq $ip) { continue }
                if (Test-HostOnline $target $TimeoutMilliseconds) { [void]$targets.Add($target) }
            }
        } catch {
            ('Failed to scan adapter IP ' + $ip + ': ' + $_.Exception.Message) | Tee-Object -FilePath $logPath -Append
        }
    }
}

$sortedTargets = @($targets | Sort-Object)
if (-not $sortedTargets) {
    'No online LAN hosts found.' | Tee-Object -FilePath $logPath -Append
    exit 1
}

$sortedTargets | Set-Content -LiteralPath $lanListPath -Encoding ASCII
('Discovered hosts: ' + $sortedTargets.Count) | Tee-Object -FilePath $logPath -Append
('Saved host list: ' + $lanListPath) | Tee-Object -FilePath $logPath -Append

& (Join-Path $PSScriptRoot 'remote_collect.ps1') -ComputerListPath $lanListPath
