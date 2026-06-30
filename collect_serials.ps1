# Script author: NaF

$ErrorActionPreference = 'SilentlyContinue'

$outDir = Join-Path $PSScriptRoot 'output'
$brandMapPath = Join-Path $PSScriptRoot 'brand_map.csv'
$configPath = Join-Path $PSScriptRoot 'config.csv'
if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

function U([string]$hex) {
    $result = ''
    foreach ($part in $hex.Split(' ')) { if ($part) { $result += [char]([Convert]::ToInt32($part, 16)) } }
    return $result
}

function Clean($value) {
    if ($null -eq $value) { return '' }
    return ([string]$value).Trim()
}

function Join-EdidChars($values) {
    if (-not $values) { return '' }
    return (-join ($values | Where-Object { $_ -ne 0 } | ForEach-Object { [char]$_ })).Trim()
}

function CsvEscape([string]$value) {
    if ($null -eq $value) { $value = '' }
    return '"' + $value.Replace('"', '""') + '"'
}

function Format-ExcelText([string]$value) {
    if ($null -eq $value) { return '' }
    return '="' + $value.Replace('"', '""') + '"'
}

function Format-ExportValue([string]$property, [string]$value) {
    $text = (Clean $value)
    if ($property -match 'Serial$' -and $text.Length -ge 10 -and $text -notmatch '^Not reported$') {
        return Format-ExcelText $text
    }
    return $text
}

function Is-BadSerial([string]$value) {
    $text = (Clean $value).ToUpper()
    if (-not $text) { return $true }
    if ($text -eq 'NOT REPORTED') { return $true }
    if ($text -match 'TO BE FILLED|O\.E\.M|OEM|DEFAULT STRING|SYSTEM SERIAL|NONE|UNKNOWN|N/A') { return $true }
    return $false
}

function Is-Enabled([string]$key) {
    if (-not $script:config.ContainsKey($key)) { return $true }
    $value = (Clean $script:config[$key]).ToLower()
    return @('1','true','yes','y','on') -contains $value
}

function Get-ActivationInfo {
    $products = @(Get-WmiObject SoftwareLicensingProduct | Where-Object { $_.PartialProductKey -and $_.Name -match 'Windows' })
    $product = $products | Select-Object -First 1
    if (-not $product) { return @{ Status = 'Unknown'; Channel = '' } }
    switch ([int]$product.LicenseStatus) {
        0 { $status = 'Unlicensed' }
        1 { $status = 'Licensed' }
        2 { $status = 'OOB Grace' }
        3 { $status = 'OOT Grace' }
        4 { $status = 'Non-Genuine Grace' }
        5 { $status = 'Notification' }
        6 { $status = 'Extended Grace' }
        default { $status = [string]$product.LicenseStatus }
    }
    $channel = ''
    if ($product.Description -match 'VOLUME_KMSCLIENT') { $channel = 'KMS Client' }
    elseif ($product.Description -match 'VOLUME_MAK') { $channel = 'MAK' }
    elseif ($product.Description -match 'OEM') { $channel = 'OEM' }
    elseif ($product.Description -match 'RETAIL') { $channel = 'Retail' }
    else { $channel = Clean $product.Description }
    return @{ Status = $status; Channel = $channel }
}

function Get-LocalAdminsText {
    $members = @()
    try {
        $group = [ADSI]'WinNT://./Administrators,group'
        $members = @($group.psbase.Invoke('Members') | ForEach-Object { $_.GetType().InvokeMember('Name', 'GetProperty', $null, $_, $null) })
    } catch {
        $members = @()
    }
    return ($members | Where-Object { $_ } | Sort-Object) -join '; '
}

function Get-BitLockerStatusText {
    $status = ''
    $manageBde = Join-Path $env:SystemRoot 'System32\manage-bde.exe'
    if (Test-Path -LiteralPath $manageBde) {
        $output = & $manageBde -status C: 2>$null
        $line = $output | Where-Object { $_ -match 'Protection Status|Conversion Status|Percentage Encrypted|Lock Status' }
        $status = (($line | ForEach-Object { (Clean $_) }) -join '; ')
    }
    if (-not $status) { $status = 'Not available' }
    return $status
}

$config = @{}
if (Test-Path -LiteralPath $configPath) {
    foreach ($item in (Import-Csv -Path $configPath)) {
        $key = Clean $item.Key
        if ($key) { $config[$key] = Clean $item.Enabled }
    }
}

$brandMap = @{}
if (Test-Path -LiteralPath $brandMapPath) {
    foreach ($item in (Import-Csv -Path $brandMapPath)) {
        $code = (Clean $item.Code).ToUpper()
        if ($code) { $brandMap[$code] = Clean $item.Name }
    }
}

function Get-BrandName([string]$code) {
    $key = (Clean $code).ToUpper()
    if ($brandMap.ContainsKey($key)) { return $brandMap[$key] }
    if ($key) { return $key }
    return 'Unknown'
}

function Get-HostBrandCode([string]$manufacturer) {
    $value = (Clean $manufacturer).ToUpper()
    if ($value -match 'LENOVO|\u8054\u60F3') { return 'LEN' }
    if ($value -match 'DELL|ALIENWARE|\u6234\u5C14|\u5916\u661F\u4EBA') { return 'DEL' }
    if ($value -match 'HEWLETT|\bHP\b|\u60E0\u666E') { return 'HWP' }
    if ($value -match 'ASUSTEK|ASUS|ROG|\u534E\u7855') { return 'ASU' }
    if ($value -match 'ACER|\u5B8F\u7881') { return 'ACR' }
    if ($value -match 'APPLE|\u82F9\u679C') { return 'APP' }
    if ($value -match 'SAMSUNG|\u4E09\u661F') { return 'SAM' }
    if ($value -match 'MICRO-STAR|\bMSI\b|\u5FAE\u661F') { return 'MSI' }
    if ($value -match 'TOSHIBA|\u4E1C\u829D') { return 'TOS' }
    if ($value -match 'FUJITSU|\u5BCC\u58EB\u901A') { return 'FUJ' }
    if ($value -match 'HUAWEI|\u534E\u4E3A') { return 'HUA' }
    if ($value -match 'HONOR|\u8363\u8000') { return 'HON' }
    if ($value -match 'XIAOMI|REDMI|\u5C0F\u7C73|\u7EA2\u7C73') { return 'XMI' }
    if ($value -match 'MICROSOFT|\u5FAE\u8F6F') { return 'MSF' }
    if ($value -match 'TONGFANG|THTF|\u6E05\u534E\u540C\u65B9') { return 'THT' }
    if ($value -match 'GREAT WALL|\bCEC\b|\u4E2D\u56FD\u957F\u57CE|\u957F\u57CE') { return 'GWC' }
    if ($value -match 'INSPUR|\u6D6A\u6F6E') { return 'INS' }
    if ($value -match 'FOUNDER|\u65B9\u6B63') { return 'FOU' }
    if ($value -match 'UNISPLENDOUR|UNIS|\u7D2B\u5149') { return 'UNI' }
    if ($value -match 'H3C|NEW H3C|\u65B0\u534E\u4E09') { return 'H3C' }
    if ($value -match 'ZTE|\u4E2D\u5174') { return 'ZTE' }
    if ($value -match 'DIGITAL CHINA|\u795E\u5DDE\u6570\u7801') { return 'DIG' }
    if ($value -match 'DAWNING|SUGON|\u4E2D\u79D1\u66D9\u5149|\u66D9\u5149') { return 'SGN' }
    if ($value -match 'KEDACOM|\u79D1\u8FBE') { return 'KDX' }
    if ($value -match 'ELITEGROUP|ECS|\u7CBE\u82F1') { return 'ECS' }
    if ($value -match 'LENOVO|\u8054\u60F3') { return 'LEN' }
    if ($value -match 'HUAWEI|\u534E\u4E3A') { return 'HUA' }
    if ($value -match 'DAHUA|\u5927\u534E') { return 'DAH' }
    if ($value -match 'QIHOO|360|\u5947\u864E') { return 'QIH' }
    if ($value -match 'QILINSEC|\u5947\u5B89\u4FE1') { return 'QIL' }
    if ($value -match 'BYD|\u6BD4\u4E9A\u8FEA') { return 'BYD' }
    if ($value -match 'PICO|\u5C0F\u7C73\u54B8') { return 'PCO' }
    if ($value -match 'TOPWAY|\u9876\u8DEF') { return 'TOW' }
    if ($value -match 'WALTON|\u5927\u6C76') { return 'WAP' }
    if ($value -match 'YIDONGHUA|\u4F0A\u52A8\u534E') { return 'YDH' }
    if ($value -match 'VASTDATA|\u8D85\u805A\u53D8') { return 'VIT' }
    if ($value -match 'WINGTECH|\u95FB\u6CF0') { return 'WYD' }
    if ($value -match 'MECHREVO|\u673A\u68B0\u9769\u547D') { return 'MCR' }
    if ($value -match 'MACHENIKE|\u673A\u68B0\u5E08') { return 'MKN' }
    if ($value -match 'HASEE|\u795E\u821F') { return 'HSE' }
    if ($value -match 'THUNDEROBOT|\u96F7\u795E') { return 'TNB' }
    if ($value -match 'COLORFUL|\u4E03\u5F69\u8679') { return 'COL' }
    if ($value -match 'PANASONIC|\u677E\u4E0B') { return 'PAR' }
    if ($value -match 'GIGABYTE|\u6280\u5609') { return 'GIG' }
    if ($value -match 'RAZER|\u96F7\u86C7') { return 'RAZ' }
    if ($value -match 'SUPER ?MICRO|SUPERMICRO') { return 'SUP' }
    if ($value -match 'QUANTA|\u5E7F\u8FBE') { return 'QUA' }
    if ($value -match 'AMAZON|AWS') { return 'AWS' }
    if ($value -match 'EVGA') { return 'EVG' }
    if ($value -match 'HTC|\u5B8F\u8FBE') { return 'HTC' }
    if ($value.Length -ge 3) { return $value.Substring(0, 3) }
    if ($value) { return $value }
    return 'Unknown'
}

function Get-OsInstallDate($os) {
    if (-not $os.InstallDate) { return '' }
    return ([System.Management.ManagementDateTimeConverter]::ToDateTime($os.InstallDate)).ToString('yyyy-MM-dd HH:mm:ss')
}

$L = @{
    Title = U '7535 8111 548C 663E 793A 5668 5E8F 5217 53F7 91C7 96C6 7ED3 679C'
    Time = U '91C7 96C6 65F6 95F4'
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
    MonitorInfo = U '663E 793A 5668 4FE1 606F'
    Monitor = U '663E 793A 5668'
    MonitorIndex = U '663E 793A 5668 5E8F 53F7'
    BrandCode = U '663E 793A 5668 54C1 724C 4EE3 7801'
    Brand = U '663E 793A 5668 54C1 724C'
    Model = U '663E 793A 5668 578B 53F7'
    MonitorSerial = U '663E 793A 5668 5E8F 5217 53F7'
    ExceptionNote = U '5F02 5E38 8BF4 660E'
    ScriptAuthor = U '811A 672C 4F5C 8005'
    ExportTxt = U '5DF2 5BFC 51FA 0020 0054 0058 0054'
    ExportCsv = U '5DF2 5BFC 51FA 0020 0043 0053 0056'
    ExportTsv = U '5DF2 5BFC 51FA 0020 0054 0053 0056'
    OutputDir = U '8F93 51FA 76EE 5F55'
    BrandCodeLabel = U '54C1 724C 4EE3 7801'
    BrandNameLabel = U '54C1 724C 540D 79F0'
    ModelNameLabel = U '578B 53F7 540D 79F0'
    SerialLabel = U '5E8F 5217 53F7'
}

$computer = Get-WmiObject Win32_ComputerSystem
$bios = Get-WmiObject Win32_BIOS
$os = Get-WmiObject Win32_OperatingSystem
$cpu = Get-WmiObject Win32_Processor | Select-Object -First 1
$disks = if (Is-Enabled 'CollectDisk') { @(Get-WmiObject Win32_DiskDrive) } else { @() }
$cDrive = if (Is-Enabled 'CollectVolume') { Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object -First 1 } else { $null }
$nics = if (Is-Enabled 'CollectNetwork') { @(Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true }) } else { @() }
$board = if (Is-Enabled 'CollectBoardMemory') { Get-WmiObject Win32_BaseBoard | Select-Object -First 1 } else { $null }
$memoryModules = if (Is-Enabled 'CollectBoardMemory') { @(Get-WmiObject Win32_PhysicalMemory) } else { @() }
$gpus = if (Is-Enabled 'CollectGpu') { @(Get-WmiObject Win32_VideoController) } else { @() }
$activation = if (Is-Enabled 'CollectActivation') { Get-ActivationInfo } else { @{ Status = ''; Channel = '' } }
$localAdmins = if (Is-Enabled 'CollectLocalAdmins') { Get-LocalAdminsText } else { '' }
$bitLockerStatus = if (Is-Enabled 'CollectBitLocker') { Get-BitLockerStatusText } else { '' }

$hostName = Clean $env:COMPUTERNAME
$manufacturer = Clean $computer.Manufacturer
$hostBrandCode = Get-HostBrandCode $manufacturer
$model = Clean $computer.Model
$hostSerial = Clean $bios.SerialNumber
if (-not $hostSerial) { $hostSerial = 'Not reported' }
$domainJoined = if ($computer.PartOfDomain) { 'Yes' } else { 'No' }
$domainOrWorkgroup = Clean $computer.Domain
$currentUser = Clean $computer.UserName
if (-not $currentUser) { $currentUser = Clean ($env:USERDOMAIN + '\' + $env:USERNAME) }
$ipAddress = (($nics | ForEach-Object { $_.IPAddress } | Where-Object { $_ -and ($_ -notmatch ':') }) -join '; ')
$macAddress = (($nics | ForEach-Object { Clean $_.MACAddress } | Where-Object { $_ }) -join '; ')
$osName = Clean $os.Caption
$osVersion = Clean $os.Version
$osArch = Clean $os.OSArchitecture
$osInstallDate = Get-OsInstallDate $os
$cpuName = Clean $cpu.Name
$memoryGb = [math]::Round(([double]$computer.TotalPhysicalMemory / 1GB), 2).ToString()
$diskInfo = (($disks | ForEach-Object { (Clean $_.Model) + ' ' + ([math]::Round(([double]$_.Size / 1GB), 0)).ToString() + 'GB' }) -join '; ')
$diskSerial = (($disks | ForEach-Object { Clean $_.SerialNumber } | Where-Object { $_ }) -join '; ')
if (-not $diskSerial) { $diskSerial = 'Not reported' }
$cDriveSizeGb = ''
$cDriveFreeGb = ''
$cDriveFreePercent = ''
if ($cDrive) {
    $cDriveSizeGb = [math]::Round(([double]$cDrive.Size / 1GB), 2).ToString()
    $cDriveFreeGb = [math]::Round(([double]$cDrive.FreeSpace / 1GB), 2).ToString()
    if ([double]$cDrive.Size -gt 0) { $cDriveFreePercent = ([math]::Round(([double]$cDrive.FreeSpace * 100 / [double]$cDrive.Size), 2)).ToString() }
}
$activationStatus = Clean $activation.Status
$licenseChannel = Clean $activation.Channel
$boardSerial = Clean $board.SerialNumber
if ((Is-Enabled 'CollectBoardMemory') -and -not $boardSerial) { $boardSerial = 'Not reported' }
$memoryModuleInfo = (($memoryModules | ForEach-Object { (Clean $_.BankLabel) + ' ' + ([math]::Round(([double]$_.Capacity / 1GB), 2)).ToString() + 'GB SN:' + (Clean $_.SerialNumber) }) -join '; ')
$gpuInfo = (($gpus | ForEach-Object { (Clean $_.Name) + ' ' + ([math]::Round(([double]$_.AdapterRAM / 1GB), 2)).ToString() + 'GB' }) -join '; ')

$hostExceptions = New-Object System.Collections.Generic.List[string]
if (Is-BadSerial $hostSerial) { $hostExceptions.Add('Host serial invalid') }
if (-not $ipAddress) { $hostExceptions.Add('IP not found') }
if (-not $macAddress) { $hostExceptions.Add('MAC not found') }
if (Is-BadSerial $diskSerial) { $hostExceptions.Add('Disk serial invalid') }
if ((Is-Enabled 'CollectVolume') -and $cDrive) {
    if (([double]$cDriveFreeGb -lt 10) -or ([double]$cDriveFreePercent -lt 10)) { $hostExceptions.Add('C drive low free space') }
}
if ((Is-Enabled 'CollectActivation') -and $activationStatus -ne 'Licensed') { $hostExceptions.Add('Windows not licensed or unknown') }
if ((Is-Enabled 'CollectBoardMemory') -and (Is-BadSerial $boardSerial)) { $hostExceptions.Add('Board serial invalid') }

$safeHost = $hostName -replace '[\\/:*?"<>|]', '_'
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$baseName = "${safeHost}_${stamp}_serials"
$txtPath = Join-Path $outDir ($baseName + '.txt')
$csvPath = Join-Path $outDir ($baseName + '.csv')
$tsvPath = Join-Path $outDir ($baseName + '.tsv')

$monitors = @(Get-WmiObject -Namespace root\wmi -Class WmiMonitorID)
if ($monitors.Count -eq 0) { $monitors = @($null) }

$rows = New-Object System.Collections.ArrayList
$index = 1
foreach ($monitor in $monitors) {
    $notes = New-Object System.Collections.Generic.List[string]
    foreach ($note in $hostExceptions) { $notes.Add($note) }
    if ($monitor) {
        $brandCode = Join-EdidChars $monitor.ManufacturerName
        $brandName = Get-BrandName $brandCode
        $monitorName = Join-EdidChars $monitor.UserFriendlyName
        $monitorSerial = Join-EdidChars $monitor.SerialNumberID
        if (-not $monitorSerial) { $monitorSerial = 'Not reported' }
        if (Is-BadSerial $monitorSerial) { $notes.Add('Monitor serial invalid') }
    } else {
        $brandCode = ''
        $brandName = 'Unknown'
        $monitorName = 'No monitor data'
        $monitorSerial = 'Not reported'
        $notes.Add('Monitor not reported')
    }

    $row = New-Object PSObject
    $row | Add-Member NoteProperty HostName $hostName
    $row | Add-Member NoteProperty HostBrand $manufacturer
    $row | Add-Member NoteProperty HostBrandCode $hostBrandCode
    $row | Add-Member NoteProperty HostModel $model
    $row | Add-Member NoteProperty HostSerial $hostSerial
    $row | Add-Member NoteProperty DomainJoined $domainJoined
    $row | Add-Member NoteProperty DomainOrWorkgroup $domainOrWorkgroup
    $row | Add-Member NoteProperty CurrentUser $currentUser
    $row | Add-Member NoteProperty IpAddress $ipAddress
    $row | Add-Member NoteProperty MacAddress $macAddress
    $row | Add-Member NoteProperty OsName $osName
    $row | Add-Member NoteProperty OsVersion $osVersion
    $row | Add-Member NoteProperty OsArch $osArch
    $row | Add-Member NoteProperty OsInstallDate $osInstallDate
    $row | Add-Member NoteProperty Cpu $cpuName
    $row | Add-Member NoteProperty MemoryGb $memoryGb
    $row | Add-Member NoteProperty DiskInfo $diskInfo
    $row | Add-Member NoteProperty DiskSerial $diskSerial
    $row | Add-Member NoteProperty CDriveSizeGb $cDriveSizeGb
    $row | Add-Member NoteProperty CDriveFreeGb $cDriveFreeGb
    $row | Add-Member NoteProperty CDriveFreePercent $cDriveFreePercent
    $row | Add-Member NoteProperty BitLockerStatus $bitLockerStatus
    $row | Add-Member NoteProperty ActivationStatus $activationStatus
    $row | Add-Member NoteProperty LicenseChannel $licenseChannel
    $row | Add-Member NoteProperty BoardSerial $boardSerial
    $row | Add-Member NoteProperty MemoryModules $memoryModuleInfo
    $row | Add-Member NoteProperty GpuInfo $gpuInfo
    $row | Add-Member NoteProperty LocalAdmins $localAdmins
    $row | Add-Member NoteProperty MonitorIndex $index
    $row | Add-Member NoteProperty MonitorBrandCode $brandCode
    $row | Add-Member NoteProperty MonitorBrand $brandName
    $row | Add-Member NoteProperty MonitorModel $monitorName
    $row | Add-Member NoteProperty MonitorSerial $monitorSerial
    $row | Add-Member NoteProperty ScriptAuthor 'NaF'
    $row | Add-Member NoteProperty CollectedAt (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    $row | Add-Member NoteProperty ExceptionNote ($notes -join '; ')
    [void]$rows.Add($row)
    $index++
}

$txt = New-Object System.Collections.Generic.List[string]
$txt.Add($L.Title)
$txt.Add('========================================')
$txt.Add($L.ScriptAuthor + ': NaF')
$txt.Add($L.Time + ':' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
$txt.Add($L.HostName + ':' + $hostName)
$txt.Add($L.HostBrand + ':' + $manufacturer)
$txt.Add($L.HostBrandCode + ':' + $hostBrandCode)
$txt.Add($L.HostModel + ':' + $model)
$txt.Add($L.HostSerial + ':' + $hostSerial)
$txt.Add($L.DomainJoined + ':' + $domainJoined)
$txt.Add($L.DomainOrWorkgroup + ':' + $domainOrWorkgroup)
$txt.Add($L.CurrentUser + ':' + $currentUser)
$txt.Add($L.IpAddress + ':' + $ipAddress)
$txt.Add($L.MacAddress + ':' + $macAddress)
$txt.Add($L.OsName + ':' + $osName)
$txt.Add($L.OsVersion + ':' + $osVersion)
$txt.Add($L.OsArch + ':' + $osArch)
$txt.Add($L.OsInstallDate + ':' + $osInstallDate)
$txt.Add($L.Cpu + ':' + $cpuName)
$txt.Add($L.MemoryGb + ':' + $memoryGb)
$txt.Add($L.DiskInfo + ':' + $diskInfo)
$txt.Add($L.DiskSerial + ':' + $diskSerial)
$txt.Add($L.CDriveSizeGb + ':' + $cDriveSizeGb)
$txt.Add($L.CDriveFreeGb + ':' + $cDriveFreeGb)
$txt.Add($L.CDriveFreePercent + ':' + $cDriveFreePercent)
$txt.Add($L.BitLockerStatus + ':' + $bitLockerStatus)
$txt.Add($L.ActivationStatus + ':' + $activationStatus)
$txt.Add($L.LicenseChannel + ':' + $licenseChannel)
$txt.Add($L.BoardSerial + ':' + $boardSerial)
$txt.Add($L.MemoryModules + ':' + $memoryModuleInfo)
$txt.Add($L.GpuInfo + ':' + $gpuInfo)
$txt.Add($L.LocalAdmins + ':' + $localAdmins)
$txt.Add('')
$txt.Add($L.MonitorInfo)
$txt.Add('========================================')
foreach ($row in $rows) {
    $txt.Add($L.Monitor + ' ' + $row.MonitorIndex)
    $txt.Add('  ' + $L.BrandCodeLabel + ':' + $row.MonitorBrandCode)
    $txt.Add('  ' + $L.BrandNameLabel + ':' + $row.MonitorBrand)
    $txt.Add('  ' + $L.ModelNameLabel + ':' + $row.MonitorModel)
    $txt.Add('  ' + $L.SerialLabel + ':' + $row.MonitorSerial)
    $txt.Add('  ' + $L.ExceptionNote + ':' + $row.ExceptionNote)
}
[System.IO.File]::WriteAllLines($txtPath, $txt.ToArray(), [System.Text.Encoding]::UTF8)

$headers = @($L.HostName,$L.HostBrand,$L.HostBrandCode,$L.HostModel,$L.HostSerial,$L.DomainJoined,$L.DomainOrWorkgroup,$L.CurrentUser,$L.IpAddress,$L.MacAddress,$L.OsName,$L.OsVersion,$L.OsArch,$L.OsInstallDate,$L.Cpu,$L.MemoryGb,$L.DiskInfo,$L.DiskSerial,$L.CDriveSizeGb,$L.CDriveFreeGb,$L.CDriveFreePercent,$L.BitLockerStatus,$L.ActivationStatus,$L.LicenseChannel,$L.BoardSerial,$L.MemoryModules,$L.GpuInfo,$L.LocalAdmins,$L.MonitorIndex,$L.BrandCode,$L.Brand,$L.Model,$L.MonitorSerial,$L.ScriptAuthor,$L.Time,$L.ExceptionNote)
$props = @('HostName','HostBrand','HostBrandCode','HostModel','HostSerial','DomainJoined','DomainOrWorkgroup','CurrentUser','IpAddress','MacAddress','OsName','OsVersion','OsArch','OsInstallDate','Cpu','MemoryGb','DiskInfo','DiskSerial','CDriveSizeGb','CDriveFreeGb','CDriveFreePercent','BitLockerStatus','ActivationStatus','LicenseChannel','BoardSerial','MemoryModules','GpuInfo','LocalAdmins','MonitorIndex','MonitorBrandCode','MonitorBrand','MonitorModel','MonitorSerial','ScriptAuthor','CollectedAt','ExceptionNote')
$csvLines = New-Object System.Collections.Generic.List[string]
$csvLines.Add(($headers | ForEach-Object { CsvEscape $_ }) -join ',')
foreach ($row in $rows) {
    $values = foreach ($prop in $props) { Format-ExportValue $prop ([string]$row.$prop) }
    $csvLines.Add(($values | ForEach-Object { CsvEscape $_ }) -join ',')
}
[System.IO.File]::WriteAllLines($csvPath, $csvLines.ToArray(), (New-Object System.Text.UTF8Encoding($true)))

$tsvLines = New-Object System.Collections.Generic.List[string]
$tsvLines.Add($headers -join "`t")
foreach ($row in $rows) {
    $values = foreach ($prop in $props) { (Format-ExportValue $prop ([string]$row.$prop)).Replace("`t", ' ') }
    $tsvLines.Add($values -join "`t")
}
[System.IO.File]::WriteAllLines($tsvPath, $tsvLines.ToArray(), [System.Text.Encoding]::Unicode)

Write-Host '========================================'
Write-Host $L.Title
Write-Host '========================================'
Write-Host ($L.HostName + ':' + $hostName)
Write-Host ($L.HostBrand + ':' + $manufacturer)
Write-Host ($L.HostBrandCode + ':' + $hostBrandCode)
Write-Host ($L.HostModel + ':' + $model)
Write-Host ($L.HostSerial + ':' + $hostSerial)
Write-Host ($L.DomainJoined + ':' + $domainJoined)
Write-Host ($L.DomainOrWorkgroup + ':' + $domainOrWorkgroup)
Write-Host ($L.CurrentUser + ':' + $currentUser)
Write-Host ($L.IpAddress + ':' + $ipAddress)
Write-Host ($L.MacAddress + ':' + $macAddress)
Write-Host ($L.OsName + ':' + $osName)
Write-Host ($L.Cpu + ':' + $cpuName)
Write-Host ($L.MemoryGb + ':' + $memoryGb)
Write-Host ($L.DiskSerial + ':' + $diskSerial)
Write-Host ($L.CDriveFreeGb + ':' + $cDriveFreeGb)
Write-Host ($L.CDriveFreePercent + ':' + $cDriveFreePercent)
Write-Host ($L.BitLockerStatus + ':' + $bitLockerStatus)
Write-Host ($L.ActivationStatus + ':' + $activationStatus)
Write-Host ($L.LicenseChannel + ':' + $licenseChannel)
Write-Host ($L.BoardSerial + ':' + $boardSerial)
Write-Host ($L.GpuInfo + ':' + $gpuInfo)
Write-Host ''
foreach ($row in $rows) {
    Write-Host ('{0} {1}: {2} / {3} / {4}: {5}' -f $L.Monitor, $row.MonitorIndex, $row.MonitorBrand, $row.MonitorModel, $L.SerialLabel, $row.MonitorSerial)
}
Write-Host ''
Write-Host ($L.ExportTxt + ':' + $txtPath)
Write-Host ($L.ExportCsv + ':' + $csvPath)
Write-Host ($L.ExportTsv + ':' + $tsvPath)
Write-Host ($L.OutputDir + ':' + $outDir)

if (Is-Enabled 'OpenOutputFolder') { Start-Process explorer.exe $outDir }
