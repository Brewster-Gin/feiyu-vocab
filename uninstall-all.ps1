# CET-6 单词计划 · 一键卸载
# 删除：计划任务、启动文件夹遗留、临时文件；可选删除整个计划文件夹
# 用法：一键卸载.cmd（会自动提权）  或  powershell -File uninstall-all.ps1
param(
    [switch]$KeepFolder,
    [switch]$Yes,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$taskName = 'CET6-Word-Reminder'

Write-Host ''
Write-Host '==== 六级单词计划 · 卸载 ====' -ForegroundColor Cyan
Write-Host ("文件夹: " + $root)
if ($DryRun) { Write-Host '（演练模式：只显示将删除什么，不会真删）' -ForegroundColor Yellow }
Write-Host ''

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host '⚠ 当前不是管理员，计划任务大概率删不掉——请双击「一键卸载.cmd」。' -ForegroundColor Yellow
    Write-Host ''
}

# ---------- 0. 关掉正在运行的弹窗 ----------
$winTitles = @('CET-6 Word Reminder', '拼写测试')
$closed = 0
foreach ($pr in (Get-Process powershell -ErrorAction SilentlyContinue)) {
    if ($winTitles -contains $pr.MainWindowTitle) {
        $closed++
        if (-not $DryRun) { $pr | Stop-Process -Force -ErrorAction SilentlyContinue }
    }
}
if ($closed -gt 0) {
    if ($DryRun) { Write-Host ('[演练] 关闭 ' + $closed + ' 个正在运行的弹窗') }
    else { Write-Host ('[OK] 已关闭 ' + $closed + ' 个正在运行的弹窗') -ForegroundColor Green }
} else {
    Write-Host '[OK] 没有正在运行的弹窗' -ForegroundColor DarkGray
}

# ---------- 1. 计划任务 ----------
$taskGone = $false
if ($DryRun) {
    Write-Host ('[演练] 删除计划任务: ' + $taskName)
    $taskGone = $true
} else {
    $ErrorActionPreference = 'Continue'
    $out = (& schtasks.exe /delete /tn $taskName /f 2>&1 | Out-String)
    $code = $LASTEXITCODE
    $ErrorActionPreference = 'Stop'
    if ($code -eq 0) {
        Write-Host ('[OK] 已删除计划任务: ' + $taskName) -ForegroundColor Green
        $taskGone = $true
    } elseif ($out -match 'cannot find|找不到|does not exist|不存在') {
        Write-Host '[OK] 计划任务本来就不存在' -ForegroundColor DarkGray
        $taskGone = $true
    } else {
        Write-Host ('[--] 计划任务未删除: ' + $out.Trim()) -ForegroundColor Yellow
    }
}

# ---------- 2. 启动文件夹遗留 ----------
$startup = [Environment]::GetFolderPath('Startup')
$leftovers = @('CET6-Word-Reminder.vbs', '六级单词提醒.vbs')
$foundStartup = $false
foreach ($f in $leftovers) {
    $p = Join-Path $startup $f
    if (Test-Path $p) {
        $foundStartup = $true
        if ($DryRun) { Write-Host ('[演练] 删除启动项: ' + $p) }
        else { Remove-Item $p -Force; Write-Host ('[OK] 已删除启动项: ' + $p) -ForegroundColor Green }
    }
}
if (-not $foundStartup) { Write-Host '[OK] 启动文件夹干净，没有 CET6 遗留项' -ForegroundColor DarkGray }

# ---------- 3. 桌面快捷方式 ----------
$desk = [Environment]::GetFolderPath('Desktop')
$lnkNames = @('肥鱼背单词.lnk', '六级背单词.lnk')
$foundLnk = $false
foreach ($n in $lnkNames) {
    $p2 = Join-Path $desk $n
    if (Test-Path $p2) {
        $foundLnk = $true
        if ($DryRun) { Write-Host ('[演练] 删除桌面图标: ' + $p2) }
        else { Remove-Item $p2 -Force; Write-Host ('[OK] 已删除桌面图标: ' + $p2) -ForegroundColor Green }
    }
}
if (-not $foundLnk) { Write-Host '[OK] 桌面上没有 CET6 图标' -ForegroundColor DarkGray }

# ---------- 3.5 开始菜单 ----------
$startMenu = Join-Path ([Environment]::GetFolderPath('Programs')) '肥鱼背单词'
$startMenuOld = Join-Path ([Environment]::GetFolderPath('Programs')) '六级背单词'
if (Test-Path $startMenuOld) { Remove-Item $startMenuOld -Recurse -Force -ErrorAction SilentlyContinue }
if (Test-Path $startMenu) {
    if ($DryRun) { Write-Host ('[演练] 删除开始菜单项: ' + $startMenu) }
    else { Remove-Item $startMenu -Recurse -Force -ErrorAction SilentlyContinue; Write-Host '[OK] 已删除开始菜单项' -ForegroundColor Green }
} else {
    Write-Host '[OK] 开始菜单没有 CET6 项' -ForegroundColor DarkGray
}

# ---------- 4. 图标目录 ----------
$iconDirs = @(
    (Join-Path $env:LOCALAPPDATA 'CET6Icon'),
    (Join-Path $env:ProgramData 'CET6Icon')
)
$foundIcon = $false
foreach ($dir in $iconDirs) {
    if (Test-Path $dir) {
        $foundIcon = $true
        if ($DryRun) { Write-Host ('[演练] 删除图标目录: ' + $dir) }
        else { Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue; Write-Host ('[OK] 已删除图标目录: ' + $dir) -ForegroundColor Green }
    }
}
if (-not $foundIcon) { Write-Host '[OK] 没有图标目录残留' -ForegroundColor DarkGray }

# ---------- 5. 临时文件 ----------
$tmp = @()
$tmp += Get-ChildItem -Path $env:TEMP -Filter 'cet6*' -ErrorAction SilentlyContinue
if ($tmp.Count -gt 0) {
    foreach ($t in $tmp) {
        if ($DryRun) { Write-Host ('[演练] 删除临时文件: ' + $t.FullName) }
        else { Remove-Item $t.FullName -Force -Recurse -ErrorAction SilentlyContinue; Write-Host ('[OK] 已删除: ' + $t.FullName) -ForegroundColor Green }
    }
} else {
    Write-Host '[OK] 没有残留临时文件' -ForegroundColor DarkGray
}

# ---------- 6. 文件夹 ----------
Write-Host ''
$deleteFolder = $false
if (-not $KeepFolder) {
    $answer = 'n'
    if ($Yes) {
        $answer = 'y'
    } elseif (-not $DryRun) {
        Write-Host '打卡表和学习记录都在这个文件夹里。' -ForegroundColor Yellow
        $answer = Read-Host '是否连文件夹一起删除？[y/N]'
    }
    if ($answer -match '^(y|Y)') { $deleteFolder = $true }
}

if ($deleteFolder) {
    if ($DryRun) {
        Write-Host ('[演练] 将删除整个文件夹: ' + $root)
    } else {
        $bat = Join-Path $env:TEMP ('cet6_cleanup_' + [guid]::NewGuid().ToString('N') + '.cmd')
        $lines = @(
            '@echo off',
            'timeout /t 2 /nobreak >nul',
            ('rd /s /q "' + $root + '"'),
            'del /f /q "%~f0"'
        )
        Set-Content -Path $bat -Value $lines -Encoding ASCII
        Start-Process -FilePath 'cmd.exe' -ArgumentList '/c', ('"' + $bat + '"') -WindowStyle Hidden
        Write-Host '[OK] 文件夹将在 2 秒后删除（本窗口会自动关闭）' -ForegroundColor Green
    }
} else {
    Write-Host '[--] 保留文件夹（打卡表、词库、记录都还在）' -ForegroundColor DarkGray
}

# ---------- 7. 复核 ----------
Write-Host ''
if (-not $DryRun) {
    Write-Host '—— 复核 ——' -ForegroundColor DarkGray
    $stillTask = (schtasks /query /tn $taskName 2>&1 | Select-String 'CET6') -ne $null
    $stillLnk = (Test-Path (Join-Path $desk '肥鱼背单词.lnk')) -or (Test-Path (Join-Path $desk '六级背单词.lnk'))
    $stillIcon = $false
    foreach ($dir in $iconDirs) { if (Test-Path $dir) { $stillIcon = $true } }
    $stillStartup = $false
    foreach ($f in $leftovers) { if (Test-Path (Join-Path $startup $f)) { $stillStartup = $true } }
    $stillMenu = Test-Path $startMenu

    Write-Host ('  计划任务: ' + $(if ($stillTask) { '仍在（需要管理员）' } else { '已清除' }))
    Write-Host ('  桌面图标: ' + $(if ($stillLnk) { '仍在' } else { '已清除' }))
    Write-Host ('  图标目录: ' + $(if ($stillIcon) { '仍在' } else { '已清除' }))
    Write-Host ('  启动项  : ' + $(if ($stillStartup) { '仍在' } else { '已清除' }))
    Write-Host ('  开始菜单: ' + $(if ($stillMenu) { '仍在' } else { '已清除' }))
    if ($stillTask -or $stillLnk -or $stillIcon -or $stillStartup -or $stillMenu) {
        Write-Host '  有项目没清掉，请用管理员权限重跑一次。' -ForegroundColor Yellow
    } else {
        Write-Host '  全部清干净了。' -ForegroundColor Green
    }
}

Write-Host ''
if ($taskGone) {
    Write-Host '卸载完成：任务计划、桌面图标、图标目录、启动项、临时文件都已清理。' -ForegroundColor Cyan
} else {
    Write-Host '计划任务没删掉，请用管理员权限重新运行「一键卸载.cmd」。' -ForegroundColor Yellow
}
if (-not $DryRun) { Write-Host ''; Read-Host '按回车关闭' | Out-Null }
