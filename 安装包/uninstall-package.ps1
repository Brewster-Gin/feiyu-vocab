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

foreach ($pr in (Get-Process powershell -ErrorAction SilentlyContinue)) {
    if ($pr.MainWindowTitle -in @('CET-6 Word Reminder', '拼写测试')) {
        $pr | Stop-Process -Force -ErrorAction SilentlyContinue
    }
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

Write-Host ''
if (-not $KeepFolder) {
    $ans = Read-Host '是否连程序文件夹一起删除（打卡表和学习记录都在里面）？[y/N]'
    if ($ans -match '^(y|Y)') {
        $bat = Join-Path $env:TEMP ('cet6_cleanup_' + [guid]::NewGuid().ToString('N') + '.cmd')
        $lines = @('@echo off', 'timeout /t 2 /nobreak >nul', ('rd /s /q "' + $dest + '"'), 'del /f /q "%~f0"')
        Set-Content -Path $bat -Value $lines -Encoding ASCII
        Start-Process -FilePath 'cmd.exe' -ArgumentList '/c', ('"' + $bat + '"') -WindowStyle Hidden
        Write-Host '[OK] 文件夹将在 2 秒后删除' -ForegroundColor Green
    } else {
        Write-Host '[--] 已保留程序文件夹' -ForegroundColor DarkGray
    }
}

Write-Host ''
Write-Host '卸载完成。' -ForegroundColor Cyan
Write-Host ''
