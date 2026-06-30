# 电脑和显示器序列号采集工具

脚本作者：NaF

## 功能

- 一键显示并导出本机信息：主机名、主机品牌、主机品牌代码、主机型号、主机序列号。
- 一键导出网络信息：IP 地址、MAC 地址。
- 一键导出域/工作组信息：是否加域、域名或工作组、当前登录用户。
- 一键导出系统和硬件信息：操作系统、系统版本、系统位数、安装日期、CPU、内存、硬盘信息、硬盘序列号。
- 一键导出 C 盘容量、剩余空间、剩余百分比，并自动标记低空间异常。
- 一键检查 BitLocker 状态，适合安全巡检。
- 一键导出 Windows 激活状态、许可证通道、主板序列号、内存条信息、显卡信息、本地管理员列表。
- 支持通过 `config.csv` 控制采集开关，避免采集项过多或远程采集太慢。
- 一键显示并导出显示器信息：品牌代码、品牌名称、型号、序列号。
- 自动翻译常见品牌代码，映射表可在 `brand_map.csv` 中维护。
- 自动标记异常，例如序列号为空、`To Be Filled By O.E.M.`、显示器未上报等。
- 文件名自动带上主机名和采集时间，方便区分多台电脑。
- 同时导出 `txt`、`csv`、`tsv` 三种格式。
- 支持把多台电脑采集结果一键汇总，默认同一主机只保留最新采集记录。

## 文件说明

| 文件 | 作用 |
| --- | --- |
| `show_serials.bat` | 本机一键采集入口，双击运行 |
| `collect_serials.ps1` | 本机采集主逻辑 |
| `merge_serials_summary.bat` | 多台电脑结果一键汇总入口，双击运行 |
| `merge_serials_summary.ps1` | 汇总主逻辑 |
| `remote_collect.bat` | 远程批量采集入口，适合域环境或已启用 WinRM 的内网 |
| `remote_collect.ps1` | 远程批量采集主逻辑 |
| `computers.txt` | 远程采集电脑名或 IP 列表 |
| `lan_collect.bat` | 局域网自动发现并批量采集入口 |
| `lan_collect.ps1` | 局域网自动发现并调用远程采集主逻辑 |
| `computers_lan.txt` | 局域网自动发现生成的临时目标列表 |
| `brand_map.csv` | 品牌代码映射表 |
| `config.csv` | 采集开关配置文件 |
| `output` | 本机采集结果输出目录 |
| `summary` | 多台电脑汇总结果输出目录 |

## 使用方法

### 1. 采集本机信息

双击运行：

```bat
show_serials.bat
```

运行后会在 `output` 文件夹生成 3 个文件：

| 格式 | 用途 |
| --- | --- |
| `.txt` | 中文友好的查看版 |
| `.csv` | UTF-8 BOM 编码，适合批量统计 |
| `.tsv` | UTF-16 编码，Excel 直接打开更不容易乱码 |

文件名示例：

```text
DESKTOP-001_20260630_165530_serials.csv
```

### 2. 汇总多台电脑信息

在每台电脑上运行 `show_serials.bat` 后，把各电脑生成的 `*_serials.csv` 文件复制到同一台电脑的 `output` 文件夹。

然后双击运行：

```bat
merge_serials_summary.bat
```

汇总结果会生成到 `summary` 文件夹：

```text
serials_summary_20260630_170000.csv
serials_summary_20260630_170000.tsv
```

汇总时会按“主机名”去重，同一主机多次采集时默认保留最新采集时间对应的记录。

### 3. 远程批量采集

把目标电脑名或 IP 写入 `computers.txt`，每行一台，`#` 开头的行会被忽略。

然后双击运行：

```bat
remote_collect.bat
```

远程采集依赖 PowerShell Remoting，适合域环境或已启用 WinRM 的内网。目标电脑需要允许远程 PowerShell，当前账号需要有远程执行权限。

采集完成后，各电脑的 CSV 会拉回本机 `output` 文件夹，再运行 `merge_serials_summary.bat` 即可汇总。

### 4. 局域网自动发现并采集

双击运行：

```bat
lan_collect.bat
```

脚本会根据本机启用网卡的 IPv4 地址和子网掩码扫描同网段在线主机，生成 `computers_lan.txt`，然后自动调用远程批量采集。

局域网采集仍然依赖 PowerShell Remoting/WinRM。能 ping 通只代表主机在线，不代表一定能远程采集；如果目标电脑未开启 WinRM、账号无权限或防火墙阻止远程 PowerShell，会在 `remote_collect.log` 中记录失败原因。

## 输出字段

| 字段 | 说明 |
| --- | --- |
| 主机名 | Windows 计算机名 |
| 主机品牌 | 电脑厂商，例如 Lenovo、Dell、HP |
| 主机品牌代码 | 根据电脑厂商自动生成的品牌代码，例如 LEN、DEL、HWP |
| 主机型号 | 电脑型号 |
| 主机序列号 | BIOS 序列号 |
| 是否加域 | `Yes` 表示已加入域，`No` 表示工作组 |
| 域或工作组 | 当前域名或工作组名称 |
| 当前登录用户 | 当前登录的 Windows 用户 |
| IP地址 | 当前启用网卡的 IPv4 地址，多个用分号分隔 |
| MAC地址 | 当前启用网卡的 MAC 地址，多个用分号分隔 |
| 操作系统 | Windows 系统名称 |
| 系统版本 | Windows 版本号 |
| 系统位数 | 32 位或 64 位 |
| 安装日期 | 操作系统安装日期 |
| CPU | 处理器型号 |
| 内存(GB) | 物理内存容量 |
| 硬盘信息 | 硬盘型号和容量 |
| 硬盘序列号 | 物理硬盘序列号，多个用分号分隔 |
| C盘容量(GB) | C 盘总容量 |
| C盘剩余(GB) | C 盘剩余空间 |
| C盘剩余百分比 | C 盘剩余空间百分比 |
| BitLocker状态 | C 盘 BitLocker/加密状态 |
| Windows激活状态 | Windows 授权状态，例如 Licensed |
| 许可证通道 | KMS Client、MAK、OEM、Retail 等 |
| 主板序列号 | 主板序列号 |
| 内存条信息 | 每根内存条的槽位、容量、序列号 |
| 显卡信息 | 显卡型号和显存 |
| 本地管理员 | 本机 Administrators 组成员 |
| 显示器序号 | 第几台显示器 |
| 显示器品牌代码 | EDID 品牌代码，例如 DEL、LEN |
| 显示器品牌 | 自动翻译后的品牌名 |
| 显示器型号 | 显示器上报的型号名称 |
| 显示器序列号 | 显示器上报的序列号 |
| 采集时间 | 信息采集时间 |
| 异常说明 | 序列号无效、IP/MAC 未找到、显示器未上报等提示 |

## 品牌映射

品牌代码维护在 `brand_map.csv`：

```csv
Code,Name,Type
DEL,Dell,Both
LEN,Lenovo,Both
```

新增品牌时直接追加一行即可。`Type` 可填写 `Host`、`Monitor` 或 `Both`，主要用于人工维护说明。

## 配置开关

采集开关维护在 `config.csv`，`Enabled` 填 `1` 表示开启，填 `0` 表示关闭。

| Key | 作用 |
| --- | --- |
| `CollectActivation` | 采集 Windows 激活状态和许可证通道 |
| `CollectBoardMemory` | 采集主板序列号和内存条信息 |
| `CollectGpu` | 采集显卡信息 |
| `CollectLocalAdmins` | 采集本地管理员列表 |
| `CollectDisk` | 采集硬盘信息和硬盘序列号 |
| `CollectVolume` | 采集 C 盘容量、剩余空间和剩余百分比 |
| `CollectBitLocker` | 采集 BitLocker 状态 |
| `CollectNetwork` | 采集 IP 和 MAC 地址 |
| `OpenOutputFolder` | 本机采集完成后自动打开输出目录 |
| `SummaryDeduplicate` | 汇总时同一主机只保留最新采集记录 |

## 兼容性

- 支持 Windows 7、Windows 10、Windows 11。
- 使用系统自带 `PowerShell` 和 `WMI`，不需要安装第三方软件。
- 远程批量采集需要 PowerShell Remoting/WinRM 支持。
- 局域网自动发现默认每个子网最多扫描 254 个地址，适合常见 `/24` 内网。
- 如果旧系统控制台显示中文乱码，以导出的 `csv` 或 `tsv` 文件为准。

## 注意事项

- 部分显示器、扩展坞、转接线可能不会上报真实显示器序列号，脚本会显示 `Not reported`。
- C 盘剩余空间低于 10GB 或低于 10% 时，会在“异常说明”中标记 `C drive low free space`。
- Windows 7 部分版本没有 BitLocker 或 `manage-bde`，脚本会显示 `Not available`。
- 如果显示器品牌代码不在内置映射表中，会直接显示原始代码。
- 建议批量统计时优先使用 `.tsv`，Excel 直接打开通常最稳定。
- 本机采集和汇总完成后会自动打开输出目录。
