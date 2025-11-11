# ======================================================================================
# 脚本名称: EscapeFromDuckovModCounter.ps1
# 功能概述（中文注释，仅注释含中文；代码字符串均为 ASCII，运行时动态拼装中文，避免编码问题）:
#   1) 访问 Steam 创意工坊页面（Escape From Duckov / 逃离雅科夫），抓取并解析该游戏的 MOD 总数量。
#   2) 将统计结果写入当前工作区下的 excel 文件夹，以“逃离雅科夫-MM月dd日-Mods统计.xlsx”命名的每日文件中，按行追加。
#   3) 生成系统通知（中文），通知格式示例：
#      “2025年11月04日，逃离雅科夫创意工坊市场的Mod总数量为333个，比昨天321个多上架了12个！”
#   4) 便于通过 Windows 计划任务设置每天 07:00 自动执行。
# 使用方式:
#   - 手动执行（当前用户）：
#       

#   - 计划任务（每天 07:00）：见本文件底部“计划任务示例”注释。
# 运行前提:
#   - Windows PowerShell 5.1 或以上版本
#   - 本机安装 Microsoft Excel（通过 COM 自动化写入 .xlsx）
# 重要说明:
#   - 为避免 PowerShell 5.1 在非 UTF-8 编码环境中的解析问题，脚本中的中文仅用于注释；运行时展示给用户的中文文本使用 Unicode 码点拼接生成。
# ======================================================================================

param()

# ========== 新增：强制启用 TLS 1.2 ==========
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[Net.ServicePointManager]::Expect100Continue = $false
[System.Net.ServicePointManager]::DefaultConnectionLimit = 10
# 控制台输出编码设为 UTF-8，避免中文在控制台显示成乱码（不影响脚本逻辑）
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
# =========================================

# -------------------------------
# 基本配置（可按需调整）
# -------------------------------
$cfg = [pscustomobject]@{
	WorkshopUrl = "https://steamcommunity.com/app/3167020/workshop/"  # 目标创意工坊页面
	UserAgent   = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0 Safari/537.36"
	TimeoutSec  = 30                                                   # HTTP 超时（秒）
	OutputDir   = "$PSScriptRoot\excel"                                # Excel 输出目录（当前工作区/excel）
	UseMirror   = $true                                                 # 受限网络建议启用：通过镜像读取公开页面
}



# -------------------------------
# 工具函数：HTTP 请求（自定义 UA/超时/语言）
# -------------------------------
function Invoke-Http {
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [int]$TimeoutSec = 30,
        [string]$UserAgent = $cfg.UserAgent
    )

    # 在受限网络（公共 WiFi 等）下可启用镜像读取公开 HTML（r.jina.ai）
    $effectiveUrl = if ($cfg.UseMirror) { "https://r.jina.ai/http://" + ($Url -replace '^https?://','') } else { $Url }

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            # 优先使用 HttpWebRequest，规避部分环境下 IWR 的握手/发送阶段问题
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            [Net.ServicePointManager]::Expect100Continue = $false

            $req = [System.Net.HttpWebRequest]::Create($effectiveUrl)
            $req.Method = 'GET'
            $req.UserAgent = $UserAgent
            $req.Accept = 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
            $req.Headers.Add('Accept-Language','en-US,en;q=0.9')
            $req.ProtocolVersion = [Version]'1.1'   # 强制 HTTP/1.1
            $req.KeepAlive = $false                 # 关闭 Keep-Alive
            $req.Timeout = $TimeoutSec * 1000
            $req.ReadWriteTimeout = $TimeoutSec * 1000
            $req.AllowAutoRedirect = $true
            $req.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate

            $resp = $req.GetResponse()
            try {
                $stream = $resp.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($stream)
                $html = $reader.ReadToEnd()
                $reader.Close()
            } finally {
                $resp.Close()
            }
            return [pscustomobject]@{ Content = $html; StatusCode = 200 }
        } catch {
            if ($attempt -ge 2) {
                break
            }
            Start-Sleep -Seconds 2
        }
    }

    # 回退方案：使用系统 curl.exe（随 Windows 10+ 附带），跟随重定向并启用压缩
    try {
        $tmp = [System.IO.Path]::GetTempFileName()
        # 使用单字符串参数并显式包裹需要空格的值，避免被拆分
        $uaEsc   = '"' + $UserAgent + '"'
        $accEsc  = '"Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"'
        $langEsc = '"Accept-Language: en-US,en;q=0.9"'
        $tmpEsc  = '"' + $tmp + '"'
        $urlEsc  = '"' + $effectiveUrl + '"'
        $argStr = "-sS -L --compressed -A $uaEsc -H $accEsc -H $langEsc --output $tmpEsc $urlEsc"
        $p = Start-Process -FilePath 'curl.exe' -ArgumentList $argStr -NoNewWindow -PassThru -Wait
        if ($p.ExitCode -ne 0) { throw "curl exit $($p.ExitCode)" }
        $html = Get-Content -Raw -Path $tmp
        Remove-Item -Force $tmp -ErrorAction SilentlyContinue
        return [pscustomobject]@{ Content = $html; StatusCode = 200 }
    } catch {
        throw
    }
}

# -------------------------------
# 工具函数：解析 MOD 总数（仅使用英文/通用结构以提升稳定性）
#   - 兼容以下常见模式：
#       “See all N Mods”
#       “Showing x-y of N entries”
#       “id=searchResults_total” 容器内的 N
# -------------------------------
function Parse-WorkshopModCount {
	param([string]$Html)

	$patterns = @(
		'See\s+all\s+([\d,\.]+)\s+Mods',
		'Showing\s+\d+(?:-\d+)?\s+of\s+([\d,\.]+)\s+entries',
		'id=["'']searchResults_total["'']>\s*([\d,\.]+)\s*<'
	)

    foreach ($p in $patterns) {
        $m = [regex]::Match($Html, $p, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($m.Success) {
            $num = ($m.Groups[1].Value -replace '[^\d]', '')
            if ($num) { return [int]$num }
        }
    }
	throw "Failed to parse MOD total from Workshop page."
}

# -------------------------------
# 工具函数：ASCII 源码中构造中文字符串（传入空格分隔的十六进制 Unicode 码点）
# -------------------------------
function CN {
	param([string]$Hexes)
	return (($Hexes -split ' ') | ForEach-Object { [char]([Convert]::ToInt32($_,16)) }) -join ''
}

# -------------------------------
# 中文片段构造：游戏名、日期、通知文案所需词语
# -------------------------------
function Get-GameNameCN {
	# 逃 离 雅 科 夫
	return CN '9003 79BB 96C5 79D1 592B'
}

function Get-DateCN {
	param([datetime]$Date = (Get-Date))
	$y = $Date.ToString('yyyy')
	$m = $Date.ToString('MM')
	$d = $Date.ToString('dd')
	$year = [char]0x5E74  # 年
	$month = [char]0x6708 # 月
	$day = [char]0x65E5   # 日
	return "$y$year$m$month$d$day"
}

# -------------------------------
# 构造固定 Excel 文件名
#   目标格式：逃离雅科夫-Mods数量统计.xlsx
# -------------------------------
function Get-ExcelPath {
    $game = Get-GameNameCN
    $cnNumStat = (CN '6570 91CF 7EDF 8BA1')  # 数量统计
    $fileName = "$game-Mods$cnNumStat.xlsx"
    if (-not (Test-Path $cfg.OutputDir)) { New-Item -ItemType Directory -Path $cfg.OutputDir | Out-Null }
    return Join-Path $cfg.OutputDir $fileName
}

# -------------------------------
# 追加一行到 Excel（使用 Excel COM，需安装 Excel）
#   - 若文件不存在则创建，并写入表头：Date / Game / ModCount
#   - 每次运行在最后一行追加一条记录
# -------------------------------
function Ensure-ExcelRow {
	param(
		[Parameter(Mandatory = $true)][string]$ExcelPath,
		[datetime]$Date = (Get-Date),
		[int]$Count
	)

	$excel = $null
	try {
		$excel = New-Object -ComObject Excel.Application
		$excel.Visible = $false
		$excel.DisplayAlerts = $false

		$wb = if (Test-Path $ExcelPath) { $excel.Workbooks.Open($ExcelPath) } else { $excel.Workbooks.Add() }
		$ws = $wb.Worksheets.Item(1); if (-not $ws) { $ws = $wb.Worksheets.Add() }
		$ws.Name = "ModCounts"

        if (-not $ws.Cells.Item(1,1).Value2) {
            $ws.Cells.Item(1,1).Value2 = "Date"
            $ws.Cells.Item(1,2).Value2 = "Game"
            $ws.Cells.Item(1,3).Value2 = "ModCount"
        }

		$lastRow = $ws.UsedRange.Rows.Count; if ($lastRow -lt 1) { $lastRow = 1 }
		$appendRow = $lastRow + 1

        # 写入日期为 OLE Automation 数字，避免 COM 类型转换问题，再设置显示格式
        $cellDate = $ws.Cells.Item($appendRow, 1)
        $cellDate.Value2 = [double]([DateTime]$Date).ToOADate()
        $ws.Cells.Item($appendRow, 2).Value2 = Get-GameNameCN
        $ws.Cells.Item($appendRow, 3).Value2 = [double]$Count

        # 设置列格式与居中：A列日期格式 yyyy/mm/dd；A1:C{appendRow} 水平/垂直居中
        try {
            $ws.Columns.Item('A').NumberFormat = "yyyy/mm/dd"
        } catch {}

        try {
            $xlCenter = -4108
            $ur = $ws.UsedRange
            if ($ur -ne $null) {
                $ur.HorizontalAlignment = $xlCenter
                $ur.VerticalAlignment = $xlCenter
            }
        } catch {}

        try { $ws.Columns.Item('A:C').AutoFit() | Out-Null } catch {}

        $dir = Split-Path -Parent $ExcelPath
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
        if (Test-Path $ExcelPath) {
            try { $wb.Save() } catch { $wb.SaveAs($ExcelPath, 51) }
        } else {
            $wb.SaveAs($ExcelPath, 51)  # 51 = xlOpenXMLWorkbook (*.xlsx)
        }
        $wb.Close($true)
	} finally {
		if ($excel) { $excel.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
		[gc]::Collect(); [gc]::WaitForPendingFinalizers()
	}
}

# -------------------------------
# 读取“昨天”的 Excel 文件获取上一天的 ModCount（若无则返回 $null）
# -------------------------------
function Get-YesterdayCount {
    $path = Get-ExcelPath
    if (-not (Test-Path $path)) { return $null }

	$excel = $null
	try {
		$excel = New-Object -ComObject Excel.Application
		$excel.Visible = $false
		$excel.DisplayAlerts = $false

        $wb = $excel.Workbooks.Open($path)
		$ws = $wb.Worksheets.Item(1)
		if (-not $ws) { $wb.Close($false); return $null }

		$lastRow = $ws.UsedRange.Rows.Count
		if ($lastRow -lt 2) { $wb.Close($false); return $null }

		$val = $ws.Cells.Item($lastRow, 3).Value2
		$wb.Close($false)
		if ($val -ne $null -and $val -ne "") { return [int]$val }
		return $null
	} catch {
		return $null
	} finally {
		if ($excel) { $excel.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
		[gc]::Collect(); [gc]::WaitForPendingFinalizers()
	}
}

# -------------------------------
# 系统通知（托盘气泡）；若失败则回退到控制台输出
# -------------------------------
function Send-Toast {
    param([string]$Title, [string]$Message)
    # 优先使用 Windows 10+ Toast（Unicode 兼容更好），失败再回退托盘气泡
    try {
        # 一次性注册 AppUserModelID 到开始菜单快捷方式，确保经典应用也能显示系统 Toast
        $appId = 'EscapeFromDuckov.ModCounter'
        $shortcut = Join-Path $env:APPDATA 'Microsoft\\Windows\\Start Menu\\Programs\\EscapeFromDuckovModCounter.lnk'
        if (-not (Test-Path $shortcut)) {
            $csharp = @"
using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Runtime.InteropServices.ComTypes;

[ComImport, Guid("00021401-0000-0000-C000-000000000046"), ClassInterface(ClassInterfaceType.None)]
class CShellLink { }

[ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("000214F9-0000-0000-C000-000000000046")]
interface IShellLinkW {
    void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cchMaxPath, out IntPtr pfd, int fFlags);
    void GetIDList(out IntPtr ppidl);
    void SetIDList(IntPtr pidl);
    void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cchMaxName);
    void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
    void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cchMaxPath);
    void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
    void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cchMaxPath);
    void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
    void GetHotkey(out short pwHotkey);
    void SetHotkey(short wHotkey);
    void GetShowCmd(out int piShowCmd);
    void SetShowCmd(int iShowCmd);
    void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cchIconPath, out int piIcon);
    void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
    void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, int dwReserved);
    void Resolve(IntPtr hwnd, int fFlags);
    void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
}

[ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99")]
interface IPropertyStore {
    uint GetCount(out uint cProps);
    uint GetAt(uint iProp, out PROPERTYKEY pkey);
    uint GetValue(ref PROPERTYKEY key, out PROPVARIANT pv);
    uint SetValue(ref PROPERTYKEY key, ref PROPVARIANT pv);
    uint Commit();
}

[StructLayout(LayoutKind.Sequential, Pack=4)]
struct PROPERTYKEY { public Guid fmtid; public uint pid; }

[StructLayout(LayoutKind.Explicit)]
struct PROPVARIANT {
    [FieldOffset(0)] public ushort vt;
    [FieldOffset(8)] public IntPtr pszVal;
}

static class PropVariantHelper {
    public static PROPVARIANT FromString(string value){
        var pv = new PROPVARIANT();
        pv.vt = 31; // VT_LPWSTR
        pv.pszVal = Marshal.StringToCoTaskMemUni(value);
        return pv;
    }
}

[ComImport, Guid("0000010b-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IPersistFile { void GetClassID(out Guid pClassID); void IsDirty(); void Load([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, uint dwMode); void Save([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, bool fRemember); void SaveCompleted([MarshalAs(UnmanagedType.LPWStr)] string pszFileName); void GetCurFile([MarshalAs(UnmanagedType.LPWStr)] out string ppszFileName);} 

public class ShortcutHelper {
    static readonly PROPERTYKEY PKEY_AppUserModel_ID = new PROPERTYKEY { fmtid = new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3"), pid = 5 };
    public static void CreateShortcut(string shortcutPath, string targetPath, string arguments, string appId) {
        var link = (IShellLinkW)new CShellLink();
        link.SetPath(targetPath);
        if (!string.IsNullOrEmpty(arguments)) link.SetArguments(arguments);
        var propStore = (IPropertyStore)link;
        var pv = PropVariantHelper.FromString(appId);
        propStore.SetValue(ref PKEY_AppUserModel_ID, ref pv);
        propStore.Commit();
        var persist = (IPersistFile)link;
        persist.Save(shortcutPath, true);
    }
}
"@
            if (-not ([System.Management.Automation.PSTypeName]'ShortcutHelper').Type) {
                Add-Type -TypeDefinition $csharp -Language CSharp -ErrorAction Stop
            }
            $null = New-Item -ItemType Directory -Path (Split-Path $shortcut) -Force
            # 目标为当前 PowerShell，可不加参数，仅用于注册 AUMID
            [ShortcutHelper]::CreateShortcut($shortcut, (Get-Command powershell).Source, $null, $appId)
        }

        $safeTitle = [System.Security.SecurityElement]::Escape($Title)
        $safeMsg   = [System.Security.SecurityElement]::Escape($Message)
        Add-Type -AssemblyName 'Windows.Data'
        Add-Type -AssemblyName 'Windows.UI'
        $xml = @"
<toast>
  <visual>
    <binding template="ToastGeneric">
      <text>$safeTitle</text>
      <text>$safeMsg</text>
    </binding>
  </visual>
</toast>
"@
        $doc = New-Object Windows.Data.Xml.Dom.XmlDocument
        $doc.LoadXml($xml)
        $toast = New-Object Windows.UI.Notifications.ToastNotification $doc
        $notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId)
        $notifier.Show($toast)
        return
    } catch {
        # 回退到托盘气泡
        try {
            Add-Type -AssemblyName System.Windows.Forms
            $ni = New-Object System.Windows.Forms.NotifyIcon
            $ni.Icon = [System.Drawing.SystemIcons]::Information
            $ni.BalloonTipTitle = $Title
            $ni.BalloonTipText  = $Message
            $ni.Visible = $true
            $ni.ShowBalloonTip(8000)
            Start-Sleep -Seconds 9
            $ni.Dispose()
        } catch {
            Write-Host "$Title - $Message"
        }
    }
}

# -------------------------------
# 主流程：
#   1) 抓取页面并解析 MOD 总数
#   2) 写入每日 Excel 文件
#   3) 读取昨天数据并对比，生成中文通知文案
#   4) 发送系统通知
# -------------------------------
try {
	# 1) 抓取与解析
	$resp = Invoke-Http -Url $cfg.WorkshopUrl -TimeoutSec $cfg.TimeoutSec -UserAgent $cfg.UserAgent
	$count = Parse-WorkshopModCount -Html $resp.Content

	# 2) 写入 Excel
    $excelPath = Get-ExcelPath
	Ensure-ExcelRow -ExcelPath $excelPath -Date (Get-Date) -Count $count

	# 3) 构造中文通知文案（使用 Unicode 码点，避免编码问题）
	$cnComma = [char]0xFF0C
	$cnExcl  = [char]0xFF01
	$cnToday = Get-DateCN -Date (Get-Date)
	$cnGame  = Get-GameNameCN
	$cnWorkshop = (CN '521B 610F 5DE5 574A')  # 创意工坊
	$cnMarket   = (CN '5E02 573A')            # 市场
	$cnOf       = (CN '7684')                  # 的
	$cnMod      = 'Mod'
	$cnTotalNumIs = (CN '603B 6570 91CF 4E3A') # 总数量为
	$cnUnit     = (CN '4E2A')                  # 个
	$cnYesterday= (CN '6BD4 6628 5929')        # 比昨天
	$cnMore     = (CN '591A 4E0A 67B6 4E86')   # 多上架了
	$cnLess     = (CN '51CF 5C11 4E86')        # 减少了

	$prefix = "$cnToday$cnComma$cnGame$cnWorkshop$cnMarket$cnOf$cnMod$cnTotalNumIs$count$cnUnit"
	$yCount = Get-YesterdayCount
	if ($yCount -ne $null) {
		$diff = $count - $yCount
		if ($diff -ge 0) {
			$msg = "$prefix$cnComma$cnYesterday$yCount$cnUnit$cnMore$diff$cnUnit$cnExcl"
		} else {
			$msg = "$prefix$cnComma$cnYesterday$yCount$cnUnit$cnLess$(-$diff)$cnUnit$cnExcl"
		}
	} else {
		$cnFirst = (CN 'FF01 9996 6B21 8BB0 5F55 FF0C 6682 65E0 6628 65E5 5BF9 6BD4 3002') # ！首次记录，暂无昨日对比。
		$msg = "$prefix$cnFirst"
	}

	# 4) 通知 + 控制台输出
	Send-Toast -Title "Steam Mod 统计完成" -Message $msg
	Write-Host $msg
	exit 0
} catch {
	$err = $_.Exception.Message
	Send-Toast -Title "Steam Mod 统计失败" -Message $err
	Write-Error $err
	exit 1
}

<#
======================================================================================
计划任务注册（每天 07:00）：

# 请以管理员身份打开 PowerShell，并将脚本路径替换为你的实际路径：
$script   = "D:\\CodeWorkSpaces\\Cursor\\efd_mods_count\\EscapeFromDuckovModCounter.ps1"
$taskName = "Daily_EscapeFromDuckovModCounter_0700"
$action   = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$script`""
$trigger  = New-ScheduledTaskTrigger -Daily -At 07:00
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Description "每日 07:00 统计 Escape From Duckov 的 MOD 数量并写入 Excel" -User "$env:UserName"

手动测试：
powershell -NoProfile -ExecutionPolicy Bypass -File "D:\\CodeWorkSpaces\\Cursor\\efd_mods_count\\EscapeFromDuckovModCounter.ps1"
======================================================================================
#>


