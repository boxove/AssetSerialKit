# Script author: NaF

$ErrorActionPreference = 'SilentlyContinue'

$inDir = Join-Path $PSScriptRoot 'output'
$outDir = Join-Path $PSScriptRoot 'summary'
$configPath = Join-Path $PSScriptRoot 'config.csv'
if (-not (Test-Path -LiteralPath $inDir)) { New-Item -ItemType Directory -Path $inDir | Out-Null }
if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

function U([string]$hex) {
    $result = ''
    foreach ($part in $hex.Split(' ')) { if ($part) { $result += [char]([Convert]::ToInt32($part, 16)) } }
    return $result
}

function CsvEscape([string]$value) {
    if ($null -eq $value) { $value = '' }
    return '"' + $value.Replace('"', '""') + '"'
}

function Format-ExcelText([string]$value) {
    if ($null -eq $value) { return '' }
    return '="' + $value.Replace('"', '""') + '"'
}

function Format-ExportValue([string]$header, [string]$value) {
    $text = ([string]$value).Trim()
    if ($header -match 'Serial$' -and $text.Length -ge 10 -and $text -notmatch '^Not reported$') {
        return Format-ExcelText $text
    }
    return $text
}

function Get-Value($row, [string]$name) {
    $prop = $row.PSObject.Properties[$name]
    if ($prop) { return [string]$prop.Value }
    return ''
}

function Is-Enabled([string]$key) {
    if (-not $script:config.ContainsKey($key)) { return $true }
    $value = ([string]$script:config[$key]).Trim().ToLower()
    return @('1','true','yes','y','on') -contains $value
}

$config = @{}
if (Test-Path -LiteralPath $configPath) {
    foreach ($item in (Import-Csv -Path $configPath)) {
        $key = ([string]$item.Key).Trim()
        if ($key) { $config[$key] = ([string]$item.Enabled).Trim() }
    }
}

$L = @{
    HostName = U '4E3B 673A 540D'
    HostBrand = U '4E3B 673A 54C1 724C'
    HostBrandCode = U '4E3B 673A 54C1 724C 4EE3 7801'
    HostModel = U '4E3B 673A 578B 53F7'
    HostSerial = U '4E3B 673A 5E8F 5217 53F7'
    DomainJoined = U '662F 5426 52A0 57DF'
    DomainOrWorkgroup = U '57DF 6216 5DE5 4F5C 7EC4'
    CurrentUser = U '5F53 524D 767B 5F55 7528 6237'
    IpAddress = U '0049 0050 5730 5740'
    MacAddress = U '004D 0041 0043 5730 5740'
    OsName = U '64CD 4F5C 7CFB 7EDF'
    OsVersion = U '7CFB 7EDF 7248 672C'
    OsArch = U '7CFB 7EDF 4F4D 6570'
    OsInstallDate = U '5B89 88C5 65E5 671F'
    Cpu = U '0043 0050 0055'
    MemoryGb = U '5185 5B58 0028 0047 0042 0029'
    DiskInfo = U '786C 76D8 4FE1 606F'
    DiskSerial = U '786C 76D8 5E8F 5217 53F7'
    CDriveSizeGb = U '0043 76D8 5BB9 91CF 0028 0047 0042 0029'
    CDriveFreeGb = U '0043 76D8 5269 4F59 0028 0047 0042 0029'
    CDriveFreePercent = U '0043 76D8 5269 4F59 767E 5206 6BD4'
    BitLockerStatus = U '0042 0069 0074 004C 006F 0063 006B 0065 0072 72B6 6001'
    ActivationStatus = U '0057 0069 006E 0064 006F 0077 0073 6FC0 6D3B 72B6 6001'
    LicenseChannel = U '8BB8 53EF 8BC1 901A 9053'
    BoardSerial = U '4E3B 677F 5E8F 5217 53F7'
    MemoryModules = U '5185 5B58 6761 4FE1 606F'
    GpuInfo = U '663E 5361 4FE1 606F'
    LocalAdmins = U '672C 5730 7BA1 7406 5458'
    MonitorIndex = U '663E 793A 5668 5E8F 53F7'
    BrandCode = U '663E 793A 5668 54C1 724C 4EE3 7801'
    Brand = U '663E 793A 5668 54C1 724C'
    Model = U '663E 793A 5668 578B 53F7'
    MonitorSerial = U '663E 793A 5668 5E8F 5217 53F7'
    ScriptAuthor = U '811A 672C 4F5C 8005'
    Time = U '91C7 96C6 65F6 95F4'
    ExceptionNote = U '5F02 5E38 8BF4 660E'
    SourceFile = U '6765 6E90 6587 4EF6'
    NoFiles1 = U '6CA1 6709 627E 5230 53EF 6C47 603B 7684 0020 0043 0053 0056 0020 6587 4EF6 3002'
    NoFiles2 = U '8BF7 5148 8FD0 884C 0020 0073 0068 006F 0077 005F 0073 0065 0072 0069 0061 006C 0073 002E 0062 0061 0074 002C 0020 6216 628A 5404 7535 8111 751F 6210 7684 0020 002A 005F 0073 0065 0072 0069 0061 006C 0073 002E 0063 0073 0076 0020 653E 5165 0020 006F 0075 0074 0070 0075 0074 0020 6587 4EF6 5939 3002'
    Done = U '5C40 57DF 7F51 591A 53F0 7535 8111 4FE1 606F 6C47 603B 5B8C 6210'
    ReadFiles = U '8BFB 53D6 6587 4EF6 6570'
    RowCount = U '6C47 603B 8BB0 5F55 6570'
    DedupCount = U '53BB 91CD 540E 8BB0 5F55 6570'
    SummaryCsv = U '6C47 603B 0020 0043 0053 0056'
    SummaryTsv = U '0045 0078 0063 0065 006C 0020 53CB 597D 0020 0054 0053 0056'
}

$files = @(Get-ChildItem -Path $inDir -Filter '*_serials.csv')
if ($files.Count -eq 0) {
    Write-Host $L.NoFiles1
    Write-Host $L.NoFiles2
    exit 1
}

$rows = New-Object System.Collections.ArrayList
foreach ($file in $files) {
    $items = Import-Csv -Path $file.FullName
    foreach ($item in $items) {
        $row = New-Object PSObject
        foreach ($prop in $item.PSObject.Properties) { $row | Add-Member NoteProperty $prop.Name $prop.Value }
        $row | Add-Member NoteProperty $L.SourceFile $file.Name -Force
        [void]$rows.Add($row)
    }
}

$dedupRows = New-Object System.Collections.ArrayList
if (Is-Enabled 'SummaryDeduplicate') {
    $latestByHost = @{}
    foreach ($row in $rows) {
        $hostName = Get-Value $row $L.HostName
        $timeText = Get-Value $row $L.Time
        $key = if ($hostName) { $hostName.ToUpper() } else { Get-Value $row $L.SourceFile }
        $time = [datetime]::MinValue
        [void][datetime]::TryParse($timeText, [ref]$time)
        if (-not $latestByHost.ContainsKey($key) -or $time -gt $latestByHost[$key].Time) {
            $latestByHost[$key] = New-Object PSObject -Property @{ Time = $time; Rows = @($row) }
        } elseif ($time -eq $latestByHost[$key].Time) {
            $latestByHost[$key].Rows += $row
        }
    }
    foreach ($entry in $latestByHost.Values) { foreach ($row in $entry.Rows) { [void]$dedupRows.Add($row) } }
} else {
    foreach ($row in $rows) { [void]$dedupRows.Add($row) }
}

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$csvPath = Join-Path $outDir ("serials_summary_${stamp}.csv")
$tsvPath = Join-Path $outDir ("serials_summary_${stamp}.tsv")
$headers = @($L.HostName,$L.ScriptAuthor,$L.HostBrand,$L.HostBrandCode,$L.HostModel,$L.HostSerial,$L.DomainJoined,$L.DomainOrWorkgroup,$L.CurrentUser,$L.IpAddress,$L.MacAddress,$L.OsName,$L.OsVersion,$L.OsArch,$L.OsInstallDate,$L.Cpu,$L.MemoryGb,$L.DiskInfo,$L.DiskSerial,$L.CDriveSizeGb,$L.CDriveFreeGb,$L.CDriveFreePercent,$L.BitLockerStatus,$L.ActivationStatus,$L.LicenseChannel,$L.BoardSerial,$L.MemoryModules,$L.GpuInfo,$L.LocalAdmins,$L.MonitorIndex,$L.BrandCode,$L.Brand,$L.Model,$L.MonitorSerial,$L.Time,$L.ExceptionNote,$L.SourceFile)

$csvLines = New-Object System.Collections.Generic.List[string]
$csvLines.Add(($headers | ForEach-Object { CsvEscape $_ }) -join ',')
foreach ($row in ($dedupRows | Sort-Object { Get-Value $_ $L.HostName }, { Get-Value $_ $L.MonitorIndex })) {
    $values = foreach ($header in $headers) { Format-ExportValue $header (Get-Value $row $header) }
    $csvLines.Add(($values | ForEach-Object { CsvEscape $_ }) -join ',')
}
[System.IO.File]::WriteAllLines($csvPath, $csvLines.ToArray(), (New-Object System.Text.UTF8Encoding($true)))

$tsvLines = New-Object System.Collections.Generic.List[string]
$tsvLines.Add($headers -join "`t")
foreach ($row in ($dedupRows | Sort-Object { Get-Value $_ $L.HostName }, { Get-Value $_ $L.MonitorIndex })) {
    $values = foreach ($header in $headers) { (Format-ExportValue $header (Get-Value $row $header)).Replace("`t", ' ') }
    $tsvLines.Add($values -join "`t")
}
[System.IO.File]::WriteAllLines($tsvPath, $tsvLines.ToArray(), [System.Text.Encoding]::Unicode)

Write-Host '========================================'
Write-Host $L.Done
Write-Host '========================================'
Write-Host ($L.ReadFiles + ':' + $files.Count)
Write-Host ($L.RowCount + ':' + $rows.Count)
Write-Host ($L.DedupCount + ':' + $dedupRows.Count)
Write-Host ($L.SummaryCsv + ':' + $csvPath)
Write-Host ($L.SummaryTsv + ':' + $tsvPath)

Start-Process explorer.exe $outDir
