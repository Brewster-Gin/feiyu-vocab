# 搬家后修复：重建桌面/开始菜单快捷方式，并重新注册计划任务
#   把本文件和项目一起搬到新位置后运行
param([switch]$DryRun)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$desk = [Environment]::GetFolderPath('Desktop')
$startMenu = Join-Path ([Environment]::GetFolderPath('Programs')) '肥鱼背单词'
$vbs = Join-Path $root 'launch.vbs'
$ico = Join-Path $env:LOCALAPPDATA 'CET6Icon\cet6.ico'

Write-Host ''
Write-Host '==== 肥鱼背单词 · 修复路径 ====' -ForegroundColor Cyan
Write-Host ('当前位置: ' + $root)
if ($DryRun) { Write-Host '（演练模式，不写任何东西）' -ForegroundColor Yellow }
Write-Host ''

if (-not (Test-Path $vbs)) {
    Write-Host '[!!] 这个文件夹里没有 launch.vbs，确认是不是项目目录？' -ForegroundColor Yellow
    Read-Host '按回车退出' | Out-Null
    exit 1
}

$ws = New-Object -ComObject WScript.Shell

function New-AppShortcut([string]$path) {
    $s = $ws.CreateShortcut($path)
    $s.TargetPath = "$env:SystemRoot\System32\wscript.exe"
    $s.Arguments = '"' + $script:vbs + '" preview'
    $s.WorkingDirectory = $script:root
    if (Test-Path $script:ico) { $s.IconLocation = $script:ico + ',0' }
    $s.Description = '肥鱼背单词'
    $s.Save()
}

# ---------- 桌面 ----------
$lnk = Join-Path $desk '肥鱼背单词.lnk'
if ($DryRun) { Write-Host ('[演练] 重建桌面图标: ' + $lnk) }
else { New-AppShortcut $lnk; Write-Host '[OK] 桌面图标已重建' -ForegroundColor Green }

# ---------- 开始菜单 ----------
if ($DryRun) {
    Write-Host ('[演练] 重建开始菜单: ' + $startMenu)
} else {
    New-Item -ItemType Directory -Path $startMenu -Force | Out-Null
    New-AppShortcut (Join-Path $startMenu '肥鱼背单词.lnk')

    $u = $ws.CreateShortcut((Join-Path $startMenu '卸载肥鱼背单词.lnk'))
    $u.TargetPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    $u.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + (Join-Path $root 'uninstall-all.ps1') + '"'
    $u.WorkingDirectory = $root
    $u.Description = '卸载肥鱼背单词'
    $u.Save()
    Write-Host '[OK] 开始菜单已重建' -ForegroundColor Green
}

# ---------- 计划任务 ----------
$installer = Join-Path $root 'install-task.ps1'
if (Test-Path $installer) {
    if ($DryRun) {
        Write-Host '[演练] 重新注册计划任务（需要管理员）'
    } else {
        Write-Host '[..] 重新注册计划任务…' -ForegroundColor DarkGray
        $ErrorActionPreference = 'Continue'
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $installer 2>&1 | Out-String | Write-Host
        $ErrorActionPreference = 'Stop'
    }
}

Write-Host ''
Write-Host '完成。随便点一下桌面或开始菜单的「肥鱼背单词」验证是否能弹窗。' -ForegroundColor Cyan
Write-Host '（如果任务是手动注册的旧路径，需要用管理员身份再跑一次本脚本）'
Write-Host ''
if (-not $DryRun) { Read-Host '按回车关闭' | Out-Null }
