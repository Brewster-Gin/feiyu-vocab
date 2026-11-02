# 肥鱼背单词 · 安装
#   双击「安装.cmd」/ 安装包 exe；可自选安装位置
param(
    [switch]$DryRun,
    [switch]$NoElevate,
    [string]$Path = ''
)

$ErrorActionPreference = 'Stop'
$src = Split-Path -Parent $MyInvocation.MyCommand.Path
$today = (Get-Date).Date
$taskName = 'CET6-Word-Reminder'
$startup = [Environment]::GetFolderPath('Startup')
$desk = [Environment]::GetFolderPath('Desktop')
$startMenu = Join-Path ([Environment]::GetFolderPath('Programs')) '肥鱼背单词'

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin -and -not $NoElevate -and -not $DryRun) {
    try {
        Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList @(
            '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath, '-NoElevate') -Wait
        exit 0
    } catch { }
}

Write-Host ''
Write-Host '==== 肥鱼背单词 · 安装 ====' -ForegroundColor Cyan
Write-Host ''

# ---------- 0. 安装位置 ----------
$defaultDest = Join-Path $env:LOCALAPPDATA 'CET6Plan'
$dest = $Path
if (-not $dest) {
    if ($DryRun) {
        $dest = $defaultDest
    } else {
        Write-Host '安装到哪里？直接回车用默认位置。' -ForegroundColor White
        Write-Host ('  默认: ' + $defaultDest) -ForegroundColor DarkGray
        $ans = Read-Host '安装路径'
        if ($ans) { $dest = $ans.Trim().Trim('"') } else { $dest = $defaultDest }
    }
}
if (-not [System.IO.Path]::IsPathRooted($dest)) { $dest = Join-Path (Get-Location) $dest }

Write-Host ''
Write-Host ('安装位置: ' + $dest) -ForegroundColor White
Write-Host ('开始日期: ' + $today.ToString('yyyy-MM-dd'))
Write-Host ('管理员  : ' + $(if ($isAdmin) { '是（含解锁/定时提醒）' } else { '否（只有开机提醒）' }))
if ($DryRun) { Write-Host '（演练模式，不写任何文件）' -ForegroundColor Yellow }
Write-Host ''

$xlsx = Get-ChildItem $src -Filter '*.xlsx' | Select-Object -First 1
if (-not $xlsx) { throw '安装包缺少打卡表 xlsx' }

# ---------- 1. 复制程序 ----------
if (-not $DryRun) {
    New-Item -ItemType Directory -Path $dest -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $dest '词库') -Force | Out-Null
    $files = @('reminder.ps1', 'spell.ps1', 'strings.json', 'words.json', 'launch.vbs', 'run-spell.vbs', 'lib-dates.ps1', 'deck-manager.ps1', 'uninstall-package.ps1')
    foreach ($f in $files) {
        $p = Join-Path $src $f
        if (Test-Path $p) { Copy-Item $p (Join-Path $dest $f) -Force }
    }
    Copy-Item $xlsx.FullName (Join-Path $dest $xlsx.Name) -Force
    Copy-Item (Join-Path $src '词库\*.csv') (Join-Path $dest '词库\') -Force -ErrorAction SilentlyContinue
    Write-Host '[OK] 程序文件已复制' -ForegroundColor Green
}

# ---------- 2. 计划起点改成今天 ----------
. (Join-Path $src 'lib-dates.ps1')
if (-not $DryRun) {
    $exam = Fix-PlanDates -Dir $dest -Start $today
    Write-Host ('[OK] 计划排到 ' + $today.ToString('yyyy-MM-dd') + ' → ' + $exam.AddDays(-1).ToString('yyyy-MM-dd') + '（40 天）') -ForegroundColor Green
}

# ---------- 3. 图标 ----------
$iconPath = $null
if (-not $DryRun) {
    foreach ($dir in @('C:\Users\Public\CET6Icon', (Join-Path $dest 'icon'))) {
        try {
            New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null
            $cand = Join-Path $dir 'cet6.ico'
            Copy-Item (Join-Path $src 'cet6.ico') $cand -Force -ErrorAction Stop
            if (Test-Path $cand) { $iconPath = $cand; break }
        } catch { }
    }
    if ($iconPath) { Write-Host '[OK] 图标已就位' -ForegroundColor Green }
    else { Write-Host '[!!] 图标没地方放，桌面图标用默认样式（不影响使用）' -ForegroundColor Yellow }
}

# ---------- 4. 快捷方式：桌面 + 开始菜单 ----------
if (-not $DryRun) {
    $vbs = Join-Path $dest 'launch.vbs'
    $ws = New-Object -ComObject WScript.Shell

    $lnk = Join-Path $desk '肥鱼背单词.lnk'
    if (Test-Path $lnk) { Remove-Item $lnk -Force }
    $sc = $ws.CreateShortcut($lnk)
    $sc.TargetPath = "$env:SystemRoot\System32\wscript.exe"
    $sc.Arguments = '"' + $vbs + '" preview'
    $sc.WorkingDirectory = $dest
    if ($iconPath -and (Test-Path $iconPath)) { $sc.IconLocation = $iconPath + ',0' }
    $sc.Description = '肥鱼背单词'
    $sc.Save()

    New-Item -ItemType Directory -Path $startMenu -Force | Out-Null
    $m1 = Join-Path $startMenu '肥鱼背单词.lnk'
    if (Test-Path $m1) { Remove-Item $m1 -Force }
    $s1 = $ws.CreateShortcut($m1)
    $s1.TargetPath = "$env:SystemRoot\System32\wscript.exe"
    $s1.Arguments = '"' + $vbs + '" preview'
    $s1.WorkingDirectory = $dest
    if ($iconPath) { $s1.IconLocation = $iconPath + ',0' }
    $s1.Description = '肥鱼背单词'
    $s1.Save()

    $m2 = Join-Path $startMenu '卸载肥鱼背单词.lnk'
    if (Test-Path $m2) { Remove-Item $m2 -Force }
    $s2 = $ws.CreateShortcut($m2)
    $s2.TargetPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    $s2.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + (Join-Path $dest 'uninstall-package.ps1') + '"'
    $s2.WorkingDirectory = $dest
    $s2.Description = '卸载肥鱼背单词'
    $s2.Save()

    Write-Host '[OK] 桌面图标 + 开始菜单（含「卸载肥鱼背单词」）已创建' -ForegroundColor Green
}

# ---------- 5. 开机自启 ----------
$vbsLauncher = Join-Path $startup 'CET6-Word-Reminder.vbs'
$taskOk = $false

if ($isAdmin -and -not $DryRun) {
    $userId = "$env:USERDOMAIN\$env:USERNAME"
    $vbs = Join-Path $dest 'launch.vbs'
    $xmlPath = Join-Path $env:TEMP 'cet6_task.xml'
    $xml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo><Author>$userId</Author><Description>CET-6 daily word reminder</Description></RegistrationInfo>
  <Triggers>
    <LogonTrigger><Enabled>true</Enabled><UserId>$userId</UserId><Delay>PT40S</Delay></LogonTrigger>
    <SessionStateChangeTrigger><Enabled>true</Enabled><UserId>$userId</UserId><StateChange>SessionUnlock</StateChange></SessionStateChangeTrigger>
    <CalendarTrigger><StartBoundary>$($today.ToString('yyyy-MM-dd'))T20:30:00</StartBoundary><Enabled>true</Enabled><ScheduleByDay><DaysInterval>1</DaysInterval></ScheduleByDay></CalendarTrigger>
  </Triggers>
  <Principals><Principal id="Author"><UserId>$userId</UserId><LogonType>InteractiveToken</LogonType><RunLevel>LeastPrivilege</RunLevel></Principal></Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings><StopOnIdleEnd>false</StopOnIdleEnd><RestartOnIdle>false</RestartOnIdle></IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled><Hidden>false</Hidden><RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun><ExecutionTimeLimit>PT2H</ExecutionTimeLimit><Priority>7</Priority>
  </Settings>
  <Actions Context="Author"><Exec><Command>C:\Windows\System32\wscript.exe</Command><Arguments>"$vbs"</Arguments><WorkingDirectory>$dest</WorkingDirectory></Exec></Actions>
</Task>
"@
    [System.IO.File]::WriteAllText($xmlPath, $xml, (New-Object System.Text.UnicodeEncoding($false, $true)))
    $ErrorActionPreference = 'Continue'
    $out = (& schtasks.exe /create /tn $taskName /xml $xmlPath /f 2>&1 | Out-String)
    $ErrorActionPreference = 'Stop'
    if ($LASTEXITCODE -eq 0) { $taskOk = $true }
}

if ($taskOk) {
    if (Test-Path $vbsLauncher) { Remove-Item $vbsLauncher -Force }
    Write-Host '[OK] 已注册提醒任务（登录 / 解锁 / 每天 20:30）' -ForegroundColor Green
} elseif (-not $DryRun) {
    $vbsContent = @"
' CET-6 word reminder: run hidden at logon
WScript.Sleep 5000
Set sh = CreateObject("WScript.Shell")
sh.Run "wscript.exe ""$dest\launch.vbs""", 0, False
"@
    [System.IO.File]::WriteAllText($vbsLauncher, $vbsContent, (New-Object System.Text.UnicodeEncoding($false, $true)))
    Write-Host '[OK] 已设置开机自启（启动文件夹；只有登录提醒）' -ForegroundColor Green
} else {
    Write-Host '[演练] 会设置开机自启' -ForegroundColor Yellow
}

Write-Host ''
Write-Host '==== 安装完成 ====' -ForegroundColor Cyan
Write-Host ('安装位置    : ' + $dest) -ForegroundColor White
Write-Host ('打卡表      : ' + (Join-Path $dest $xlsx.Name))
Write-Host ('桌面图标    : 肥鱼背单词')
Write-Host ('开始菜单    : 所有程序 → 肥鱼背单词 → 卸载肥鱼背单词')
Write-Host ('卸载        : ' + (Join-Path $dest 'uninstall-package.ps1'))
Write-Host ''
if (-not $DryRun) {
    $open = Read-Host '现在打开安装文件夹看吗？[y/N]'
    if ($open -match '^(y|Y)') { Start-Process explorer.exe $dest }
}
Write-Host ''
