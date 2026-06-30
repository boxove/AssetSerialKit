# Script author: NaF

param(
    [string]$ComputerListPath = (Join-Path $PSScriptRoot 'computers.txt')
)

$ErrorActionPreference = 'Continue'

$computerListPath = $ComputerListPath
$localOutput = Join-Path $PSScriptRoot 'output'
$logPath = Join-Path $PSScriptRoot 'remote_collect.log'
$remoteDir = 'C:\Windows\Temp\serial_collect_tool'

if (-not (Test-Path -LiteralPath $localOutput)) { New-Item -ItemType Directory -Path $localOutput | Out-Null }
if (-not (Test-Path -LiteralPath $computerListPath)) {
    'Please create computers.txt first.' | Tee-Object -FilePath $logPath -Append
    exit 1
}

$computers = Get-Content -LiteralPath $computerListPath | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') }
if (-not $computers) {
    'No computer names found in computers.txt.' | Tee-Object -FilePath $logPath -Append
    exit 1
}

$sourceFiles = @(
    (Join-Path $PSScriptRoot 'collect_serials.ps1'),
    (Join-Path $PSScriptRoot 'brand_map.csv'),
    (Join-Path $PSScriptRoot 'config.csv')
) | Where-Object { Test-Path -LiteralPath $_ }

foreach ($computer in $computers) {
    $prefix = '[' + $computer + '] '
    try {
        ($prefix + 'Connecting...') | Tee-Object -FilePath $logPath -Append
        $session = New-PSSession -ComputerName $computer -ErrorAction Stop
        Invoke-Command -Session $session -ScriptBlock {
            param($remoteDir)
            if (-not (Test-Path -LiteralPath $remoteDir)) { New-Item -ItemType Directory -Path $remoteDir | Out-Null }
        } -ArgumentList $remoteDir -ErrorAction Stop

        foreach ($file in $sourceFiles) {
            Copy-Item -LiteralPath $file -Destination $remoteDir -ToSession $session -Force -ErrorAction Stop
        }

        Invoke-Command -Session $session -ScriptBlock {
            param($remoteDir)
            powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $remoteDir 'collect_serials.ps1') | Out-Null
            Get-ChildItem -Path (Join-Path $remoteDir 'output') -Filter '*_serials.csv' | Sort-Object LastWriteTime -Descending | Select-Object -First 1 | ForEach-Object { $_.FullName }
        } -ArgumentList $remoteDir -ErrorAction Stop | ForEach-Object {
            if ($_) {
                Copy-Item -FromSession $session -LiteralPath $_ -Destination $localOutput -Force -ErrorAction Stop
                ($prefix + 'Collected: ' + (Split-Path $_ -Leaf)) | Tee-Object -FilePath $logPath -Append
            }
        }

        Remove-PSSession $session
    } catch {
        ($prefix + 'Failed: ' + $_.Exception.Message) | Tee-Object -FilePath $logPath -Append
        if ($session) { Remove-PSSession $session }
    }
}

'Remote collection finished. Run merge_serials_summary.bat to build the summary.' | Tee-Object -FilePath $logPath -Append
Start-Process explorer.exe $localOutput
