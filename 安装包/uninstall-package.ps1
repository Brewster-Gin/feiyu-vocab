# 肥鱼背单词 · 卸载
#   会自动定位自己的安装目录（不依赖固定路径）
param([switch]$KeepFolder, [switch]$NoElevate)

$ErrorActionPreference = 'Stop'
$dest = Split-Path -Parent $MyInvocation.MyCommand.Path
$taskName = 'CET6-Word-Reminder'
$startup = [Environment]::GetFolderPath('Startup')
$desk = [Environment]::GetFolderPath('Desktop')
$startMenu = Join-Path ([Environment]::GetFolderPath('Programs')) '肥鱼背单词'

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin -and -not $NoElevate) {
    try {
        Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList @(
            '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath, '-NoElevate') -Wait
        exit 0
    } catch {
        Write-Host '（没有管理员权限，跳过计划任务，其余照常清理）' -ForegroundColor Yellow
    }
}

Write-Host ''
Write-Host '==== 肥鱼背单词 · 卸载 ====' -ForegroundColor Cyan
Write-Host ('程序目录: ' + $dest)
Write-Host ''

$killed = 0
foreach ($pr in (Get-Process powershell -ErrorAction SilentlyContinue)) {
    if ($pr.MainWindowTitle -in @('CET-6 Word Reminder', '拼写测试')) {
        $pr | Stop-Process -Force -ErrorAction SilentlyContinue
        $killed++
    }
}
# 隐藏运行的实例（没有窗口标题）按命令行匹配安装目录
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

$ErrorActionPreference = 'Continue'
$out = (& schtasks.exe /delete /tn $taskName /f 2>&1 | Out-String)
$code = $LASTEXITCODE
$ErrorActionPreference = 'Stop'
if ($code -eq 0) { Write-Host '[OK] 已删除计划任务' -ForegroundColor Green }
elseif ($out -match 'cannot find|找不到|does not exist|不存在') { Write-Host '[OK] 计划任务本来就不存在' -ForegroundColor DarkGray }
else { Write-Host ('[--] 计划任务未删除：' + $out.Trim()) -ForegroundColor Yellow }

foreach ($f in @('CET6-Word-Reminder.vbs', '六级单词提醒.vbs')) {
    $p = Join-Path $startup $f
    if (Test-Path $p) { Remove-Item $p -Force; Write-Host ('[OK] 已删除开机自启: ' + $f) -ForegroundColor Green }
}

foreach ($n in @('肥鱼背单词.lnk', '六级背单词.lnk')) {
    $lnk = Join-Path $desk $n
    if (Test-Path $lnk) { Remove-Item $lnk -Force; Write-Host ('[OK] 已删除桌面图标: ' + $n) -ForegroundColor Green }
}

if (Test-Path $startMenu) {
    Remove-Item $startMenu -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host '[OK] 已删除开始菜单项' -ForegroundColor Green
}

foreach ($dir in @('C:\Users\Public\CET6Icon', (Join-Path $env:LOCALAPPDATA 'CET6Icon'), (Join-Path $dest 'icon'))) {
    if (Test-Path $dir) {
        Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host ('[OK] 已删除图标目录: ' + $dir) -ForegroundColor Green
    }
}

# 移除卸载注册表项
try {
    $regKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\FeiyuVocab'
    if (Test-Path $regKey) {
        Remove-Item $regKey -Recurse -Force
        Write-Host '[OK] 已从「已安装的应用」列表移除' -ForegroundColor Green
    }
} catch { }

Write-Host ''
if (-not $KeepFolder) {
    $ans = Read-Host '是否连程序文件夹一起删除（打卡表和学习记录都在里面）？[y/N]'
    if ($ans -match '^(y|Y)') {
        $bat = Join-Path $env:TEMP ('cet6_cleanup_' + [guid]::NewGuid().ToString('N') + '.cmd')
        $lines = @('@echo off', 'timeout /t 2 /nobreak >nul', ('rd /s /q "' + $dest + '"'), 'del /f /q "%~f0"')
        Set-Content -Path $bat -Value $lines -Encoding ASCII
        Start-Process -FilePath 'cmd.exe' -ArgumentList '/c', ('"' + $bat + '"') -WindowStyle Hidden
        Start-Sleep -Milliseconds 2500

        if (Test-Path $dest) {
            # 还被占用：登记到下次登录时再删（不用重启，注销登录即可）
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
    } else {
        Write-Host '[--] 已保留程序文件夹' -ForegroundColor DarkGray
    }
}

Write-Host ''
Write-Host '卸载完成。' -ForegroundColor Cyan
Write-Host ''
