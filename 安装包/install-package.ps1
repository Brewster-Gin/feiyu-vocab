# 肥鱼背单词 · 安装
#   双击「安装.cmd」/ 安装包 exe；可自选安装位置
param(
    [switch]$DryRun,
    [switch]$NoElevate,
    [switch]$Silent,
    [switch]$NoDesktop,
    [switch]$NoStartMenu,
    [switch]$NoAutoStart,
    [string]$Path = ''
)

$ErrorActionPreference = 'Stop'
if ($Silent) { try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false } catch { } }
$src = Split-Path -Parent $MyInvocation.MyCommand.Path
$today = (Get-Date).Date
$taskName = 'CET6-Word-Reminder'
$startup = [Environment]::GetFolderPath('Startup')
$desk = [Environment]::GetFolderPath('Desktop')
$startMenu = Join-Path ([Environment]::GetFolderPath('Programs')) '肥鱼背单词'

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin -and -not $NoElevate -and -not $DryRun) {
    $fwd = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $PSCommandPath + '"'))
    if ($Path) { $fwd += @('-Path', ('"' + $Path + '"')) }
    foreach ($sw in @('Silent', 'NoDesktop', 'NoStartMenu', 'NoAutoStart')) {
        if (Get-Variable $sw -ValueOnly) { $fwd += ('-' + $sw) }
    }
    $fwd += '-NoElevate'
    try {
        Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList $fwd -WindowStyle Hidden -Wait
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

if (-not $DryRun) { $isUpgrade = Test-Path (Join-Path $dest 'words.json') } else { $isUpgrade = $false }

# ---------- 1. 复制程序 ----------
if (-not $DryRun) {
    New-Item -ItemType Directory -Path $dest -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $dest '词库') -Force | Out-Null
    $files = @('reminder.ps1', 'spell.ps1', 'strings.json', 'words.json', 'launch.vbs', 'run-spell.vbs', 'lib-dates.ps1', 'deck-manager.ps1', 'uninstall-package.ps1', 'uninstall-gui.ps1', 'uninstall-launch.vbs', 'setup-gui.ps1', 'setup-launch.vbs')
    New-Item -ItemType Directory -Path (Join-Path $dest '工具') -Force | Out-Null
    # 兜底：压缩包对子目录处理不稳，生成器若没到 工具\，从根目录补
    if (-not (Test-Path (Join-Path $dest '工具\build_plan.py'))) {
        foreach ($fn in @('build_plan.py', 'extra_words.py', 'ipa.json')) {
            $srcFile = Join-Path $src $fn
            if (Test-Path $srcFile) { Copy-Item $srcFile (Join-Path $dest ('工具\' + $fn)) -Force }
        }
    }
    if (Test-Path (Join-Path $dest '工具\build_plan.py')) {
        Write-Host '[OK] 生成器已就位（支持切换词库/改每日词数）' -ForegroundColor Green
    } else {
        Write-Host '[!!] 生成器缺失：切换词库与改每日词数将不可用' -ForegroundColor Yellow
    }
    $keepOnUpgrade = @('words.json', 'config.json')
    foreach ($f in $files) {
        $p = Join-Path $src $f
        if (-not (Test-Path $p)) { continue }
        if ($isUpgrade -and ($keepOnUpgrade -contains $f)) {
            Write-Host ('[--] 保留已有 ' + $f + '（不覆盖）') -ForegroundColor DarkGray
            continue
        }
        Copy-Item $p (Join-Path $dest $f) -Force
    }
    Copy-Item $xlsx.FullName (Join-Path $dest $xlsx.Name) -Force
    # 词库与生成器：安装包是扁平结构（IExpress 不支持子目录），这里归位
    $deckSrc = Get-ChildItem $src -Filter '*.csv' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($deckSrc) {
        $deckDir = Join-Path $dest '词库'
        $deckCount = @(Get-ChildItem $deckDir -Filter '*.csv' -ErrorAction SilentlyContinue).Count
        if ($deckCount -gt 0) {
            Write-Host ('[--] 已有 ' + $deckCount + ' 套词库，不覆盖') -ForegroundColor DarkGray
        } else {
            Copy-Item $deckSrc.FullName (Join-Path $deckDir $deckSrc.Name) -Force
            Write-Host ('[OK] 已放入默认词库：' + $deckSrc.BaseName) -ForegroundColor Green
        }
    }

    $toolNames = @('build_plan.py', 'extra_words.py', 'ipa.json')
    foreach ($fn in $toolNames) {
        $p = Join-Path $src $fn
        if (Test-Path $p) { Copy-Item $p (Join-Path $dest ('工具' + [IO.Path]::DirectorySeparatorChar + $fn)) -Force }
    }

    $genOk = Test-Path (Join-Path $dest ('工具' + [IO.Path]::DirectorySeparatorChar + 'build_plan.py'))
    if ($genOk) { Write-Host '[OK] 生成器已就位（可切换词库/改每日词数）' -ForegroundColor Green }
    else { Write-Host '[!!] 生成器缺失：切换词库与改每日词数将不可用' -ForegroundColor Yellow }
    $deckCount = @(Get-ChildItem (Join-Path $dest '词库') -Filter '*.csv' -ErrorAction SilentlyContinue).Count
    if ($deckCount -gt 0) { Write-Host ('[OK] 词库已就位（' + $deckCount + ' 套）') -ForegroundColor Green }
    else { Write-Host '[!!] 词库没拷进去，安装后需手动添加' -ForegroundColor Yellow }

    Write-Host '[OK] 程序文件已复制' -ForegroundColor Green
}

# ---------- 2. 计划起点 ----------
. (Join-Path $src 'lib-dates.ps1')
if ($isUpgrade) {
    Write-Host '[OK] 检测到已安装：保留原计划起点与打卡记录（升级安装）' -ForegroundColor Green
} elseif (-not $DryRun) {
    $exam = Fix-PlanDates -Dir $dest -Start $today
    $days = 0
    try { $days = (Get-Content (Join-Path $dest 'words.json') -Raw -Encoding UTF8 | ConvertFrom-Json).days.Count } catch { }
    Write-Host ('[OK] 计划排到 ' + $today.ToString('yyyy-MM-dd') + ' → ' + $exam.AddDays(-1).ToString('yyyy-MM-dd') + '（' + $days + ' 天）') -ForegroundColor Green
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

    if (-not $NoDesktop) {
        $lnk = Join-Path $desk '肥鱼背单词.lnk'
        if (Test-Path $lnk) { Remove-Item $lnk -Force }
        $sc = $ws.CreateShortcut($lnk)
        $sc.TargetPath = "$env:SystemRoot\System32\wscript.exe"
        $sc.Arguments = '"' + $vbs + '" preview'
        $sc.WorkingDirectory = $dest
        if ($iconPath -and (Test-Path $iconPath)) { $sc.IconLocation = $iconPath + ',0' }
        $sc.Description = '肥鱼背单词'
        $sc.Save()
    }

    if (-not $NoStartMenu) {
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
    $s2.TargetPath = 'C:\Windows\System32\wscript.exe'
    $s2.Arguments = '"' + (Join-Path $dest 'uninstall-launch.vbs') + '"'
    $s2.WorkingDirectory = $dest
    $s2.Description = '卸载肥鱼背单词'
    $s2.Save()
    }

    Write-Host '[OK] 快捷方式已创建' -ForegroundColor Green
}

# ---------- 5. 开机自启 ----------
$vbsLauncher = Join-Path $startup 'CET6-Word-Reminder.vbs'
$taskOk = $false

if ($isAdmin -and -not $DryRun -and -not $NoAutoStart) {
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
} elseif (-not $DryRun -and -not $NoAutoStart) {
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

# ---------- 6. 注册到 Windows 卸载列表 ----------
if (-not $DryRun) {
    try {
        $regKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\FeiyuVocab'
        New-Item -Path $regKey -Force | Out-Null
        New-ItemProperty -Path $regKey -Name 'DisplayName' -Value '肥鱼背单词' -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $regKey -Name 'DisplayVersion' -Value '1.3' -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $regKey -Name 'Publisher' -Value 'Brewster-Gin' -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $regKey -Name 'InstallLocation' -Value $dest -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $regKey -Name 'NoModify' -Value 1 -PropertyType DWord -Force | Out-Null
        New-ItemProperty -Path $regKey -Name 'NoRepair' -Value 1 -PropertyType DWord -Force | Out-Null
        $kb = [int](((Get-ChildItem $dest -Recurse -File -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum) / 1KB)
        New-ItemProperty -Path $regKey -Name 'EstimatedSize' -Value $kb -PropertyType DWord -Force | Out-Null
        if ($iconPath -and (Test-Path $iconPath)) {
            New-ItemProperty -Path $regKey -Name 'DisplayIcon' -Value ($iconPath + ',0') -PropertyType String -Force | Out-Null
        }
        $unCmd = '"C:\Windows\System32\wscript.exe" "' + (Join-Path $dest 'uninstall-launch.vbs') + '"'
        New-ItemProperty -Path $regKey -Name 'UninstallString' -Value $unCmd -PropertyType String -Force | Out-Null
        Write-Host '[OK] 已登记到「设置 → 已安装的应用」' -ForegroundColor Green
    } catch {
        Write-Host ('[!!] 注册卸载项失败（不影响使用）：' + $_.Exception.Message) -ForegroundColor Yellow
    }
}

Write-Host ''
Write-Host '==== 安装完成 ====' -ForegroundColor Cyan
Write-Host ('安装位置    : ' + $dest) -ForegroundColor White
Write-Host ('打卡表      : ' + (Join-Path $dest $xlsx.Name))
Write-Host ('桌面图标    : 肥鱼背单词')
Write-Host ('开始菜单    : 所有程序 → 肥鱼背单词 → 卸载肥鱼背单词')
Write-Host '卸载        : 开始菜单「卸载肥鱼背单词」（图形界面）'
Write-Host ''
if (-not $DryRun -and -not $Silent) {
    $open = Read-Host '现在打开安装文件夹看吗？[y/N]'
    if ($open -match '^(y|Y)') { Start-Process explorer.exe $dest }
}
Write-Host ''
