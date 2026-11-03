# 肥鱼背单词 · 卸载
#   会自动定位自己的安装目录（不依赖固定路径）
#   -KeepData  : 只删程序，保留打卡表 / 词库 / 学习记录
#   -Silent    : 不交互（供图形界面调用）
param([switch]$KeepData, [switch]$KeepFolder, [switch]$Silent, [switch]$NoElevate)

$ErrorActionPreference = 'Stop'
if ($Silent) { try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false } catch { } }

$dest = Split-Path -Parent $MyInvocation.MyCommand.Path
$taskName = 'CET6-Word-Reminder'
$startup = [Environment]::GetFolderPath('Startup')
$desk = [Environment]::GetFolderPath('Desktop')
$startMenu = Join-Path ([Environment]::GetFolderPath('Programs')) '肥鱼背单词'

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin -and -not $NoElevate) {
    $fwd = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $PSCommandPath + '"'))
    foreach ($sw in @('KeepData', 'KeepFolder', 'Silent')) {
        if (Get-Variable $sw -ValueOnly) { $fwd += ('-' + $sw) }
    }
    $fwd += '-NoElevate'
    try {
        Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList $fwd -Wait
        exit 0
    } catch {
        Write-Host '（没有管理员权限，跳过计划任务，其余照常清理）' -ForegroundColor Yellow
    }
}

Write-Host ''
Write-Host '==== 肥鱼背单词 · 卸载 ====' -ForegroundColor Cyan
Write-Host ('程序目录: ' + $dest)
Write-Host ''

# ---------- 1. 关掉正在运行的弹窗 ----------
$killed = 0
foreach ($pr in (Get-Process powershell -ErrorAction SilentlyContinue)) {
    if ($pr.MainWindowTitle -in @('CET-6 Word Reminder', '拼写测试')) {
        $pr | Stop-Process -Force -ErrorAction SilentlyContinue
        $killed++
    }
}
try {
    $procs = Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='wscript.exe'" -ErrorAction Stop
    foreach ($pr in $procs) {
        if ($pr.ProcessId -eq $PID) { continue }
        if ($pr.CommandLine -and $pr.CommandLine.IndexOf($dest, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            Stop-Process -Id $pr.ProcessId -Force -ErrorAction SilentlyContinue
            $killed++
        }
    }
} catch { }
if ($killed -gt 0) {
    Write-Host ('[OK] 已结束 ' + $killed + ' 个占用进程') -ForegroundColor Green
    Start-Sleep -Milliseconds 600
}

# ---------- 2. 计划任务 ----------
$ErrorActionPreference = 'Continue'
$out = (& schtasks.exe /delete /tn $taskName /f 2>&1 | Out-String)
$code = $LASTEXITCODE
$ErrorActionPreference = 'Stop'
if ($code -eq 0) { Write-Host '[OK] 已删除计划任务' -ForegroundColor Green }
elseif ($out -match 'cannot find|找不到|does not exist|不存在') { Write-Host '[OK] 计划任务本来就不存在' -ForegroundColor DarkGray }
else { Write-Host ('[--] 计划任务未删除：' + $out.Trim()) -ForegroundColor Yellow }

# ---------- 3. 开机自启 ----------
foreach ($f in @('CET6-Word-Reminder.vbs', '六级单词提醒.vbs')) {
    $p = Join-Path $startup $f
    if (Test-Path $p) { Remove-Item $p -Force; Write-Host ('[OK] 已删除开机自启: ' + $f) -ForegroundColor Green }
}

# ---------- 4. 快捷方式（只删指向本目录的） ----------
function Test-LnkMine([string]$lnkPath) {
    try {
        $sh = New-Object -ComObject WScript.Shell
        $sc = $sh.CreateShortcut($lnkPath)
        $txt = [string]$sc.TargetPath + ' ' + [string]$sc.Arguments
        return ($txt.IndexOf($dest, [StringComparison]::OrdinalIgnoreCase) -ge 0)
    } catch { return $false }
}

foreach ($n in @('肥鱼背单词.lnk', '六级背单词.lnk')) {
    $lnk = Join-Path $desk $n
    if ((Test-Path $lnk) -and (Test-LnkMine $lnk)) {
        Remove-Item $lnk -Force
        Write-Host ('[OK] 已删除桌面图标: ' + $n) -ForegroundColor Green
    }
}

if (Test-Path $startMenu) {
    $removed = 0
    foreach ($lnk in (Get-ChildItem $startMenu -Filter '*.lnk' -ErrorAction SilentlyContinue)) {
        if (Test-LnkMine $lnk.FullName) { Remove-Item $lnk.FullName -Force; $removed++ }
    }
    if ($removed -gt 0) {
        if (@(Get-ChildItem $startMenu -ErrorAction SilentlyContinue).Count -eq 0) {
            Remove-Item $startMenu -Recurse -Force -ErrorAction SilentlyContinue
        }
        Write-Host ('[OK] 已删除开始菜单项（' + $removed + ' 个）') -ForegroundColor Green
    }
}

foreach ($dir in @((Join-Path $dest 'icon'))) {
    if (Test-Path $dir) {
        Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host ('[OK] 已删除图标目录: ' + $dir) -ForegroundColor Green
    }
}

# ---------- 5. 移出「已安装的应用」 ----------
try {
    $regKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\FeiyuVocab'
    if (Test-Path $regKey) {
        $loc = [string](Get-ItemProperty $regKey -Name 'InstallLocation' -ErrorAction SilentlyContinue).InstallLocation
        if ($loc.TrimEnd('\') -eq $dest.TrimEnd('\')) {
            Remove-Item $regKey -Recurse -Force
            Write-Host '[OK] 已从「已安装的应用」列表移除' -ForegroundColor Green
        } else {
            Write-Host ('[--] 注册表项属于另一个安装（' + $loc + '），未删除') -ForegroundColor DarkGray
        }
    }
} catch { }

# ---------- 6. 程序文件 / 数据 ----------
function Remove-ProgramFiles {
    $progFiles = @(
        'reminder.ps1', 'spell.ps1', 'deck-manager.ps1', 'lib-dates.ps1', 'strings.json',
        'launch.vbs', 'run-spell.vbs', 'install-package.ps1', 'uninstall-package.ps1',
        'setup-gui.ps1', 'setup-launch.vbs', 'uninstall-gui.ps1', 'uninstall-launch.vbs',
        '安装.cmd', '卸载.cmd', 'cet6.ico'
    )
    $n = 0
    foreach ($f in $progFiles) {
        $p = Join-Path $dest $f
        if (Test-Path $p) {
            try { Remove-Item $p -Force -ErrorAction Stop; $n++ } catch { }
        }
    }
    foreach ($d in @('工具', '__pycache__')) {
        $p = Join-Path $dest $d
        if (Test-Path $p) { Remove-Item $p -Recurse -Force -ErrorAction SilentlyContinue; $n++ }
    }
    Write-Host ('[OK] 已删除 ' + $n + ' 项程序文件') -ForegroundColor Green
    Write-Host '[OK] 已保留：打卡表 xlsx、词库、学习记录、计划设置' -ForegroundColor Green
}

function Remove-Folder {
    $bat = Join-Path $env:TEMP ('cet6_cleanup_' + [guid]::NewGuid().ToString('N') + '.cmd')
    $lines = @('@echo off', 'timeout /t 2 /nobreak >nul', ('rd /s /q "' + $dest + '"'), 'del /f /q "%~f0"')
    Set-Content -Path $bat -Value $lines -Encoding ASCII
    Start-Process -FilePath 'cmd.exe' -ArgumentList '/c', ('"' + $bat + '"') -WindowStyle Hidden
    Start-Sleep -Milliseconds 2500

    if (Test-Path $dest) {
        try {
            $ro = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce'
            New-Item -Path $ro -Force | Out-Null
            New-ItemProperty -Path $ro -Name 'FeiyuCleanup' -PropertyType String -Force `
                -Value ('cmd.exe /c rd /s /q "' + $dest + '"') | Out-Null
            Write-Host '[OK] 文件夹仍被占用，已登记为下次登录时自动清理' -ForegroundColor Green
            Write-Host ('     位置: ' + $dest)
        } catch {
            Write-Host ('[--] 文件夹删不掉，请手动删除: ' + $dest) -ForegroundColor Yellow
        }
    } else {
        Write-Host '[OK] 文件夹已删除' -ForegroundColor Green
    }
}

if ($Silent) {
    if ($KeepData) { Remove-ProgramFiles } else { Remove-Folder }
} elseif ($KeepFolder) {
    Write-Host '[--] 已保留程序文件夹' -ForegroundColor DarkGray
} else {
    Write-Host ''
    Write-Host '选择卸载方式：' -ForegroundColor White
    Write-Host '  1 = 只删程序，保留打卡表和学习记录' -ForegroundColor DarkGray
    Write-Host '  2 = 全部删除（含打卡表）' -ForegroundColor DarkGray
    Write-Host '  回车 = 保留文件夹，什么都不删' -ForegroundColor DarkGray
    $ans = Read-Host '输入 1 / 2 / 回车'
    switch ($ans.Trim()) {
        '1' { Remove-ProgramFiles }
        '2' {
            $c = Read-Host '全部删除不可恢复，确认？[y/N]'
            if ($c -match '^(y|Y)') { Remove-Folder } else { Write-Host '[--] 已取消' -ForegroundColor DarkGray }
        }
        default { Write-Host '[--] 已保留程序文件夹' -ForegroundColor DarkGray }
    }
}

Write-Host ''
Write-Host '卸载完成。' -ForegroundColor Cyan
Write-Host ''
