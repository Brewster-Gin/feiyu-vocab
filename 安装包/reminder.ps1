# CET-6 daily word reminder popup (Windows, PowerShell 5.1 + WPF)
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File "reminder.ps1"
#        -Preview  force show (ignore today's state)
#        -Test     auto screenshot for self-check
#        -AsOf 2026-11-16   preview a specific date (testing only)
#        -Sync     only push existing state files into the workbook, then exit
# NOTE: this file is intentionally ASCII-only; all Chinese UI text lives in strings.json
param(
    [switch]$Preview,
    [switch]$Test,
    [switch]$Sync,
    [string]$AsOf = ''
)

$ErrorActionPreference = 'Stop'

# single-instance guard: logon / unlock / 20:30 triggers can fire together
$mutexOk = $false
try {
    $mtx = New-Object System.Threading.Mutex($false, 'Local\CET6WordReminder')
    $mutexOk = $mtx.WaitOne(0)
} catch [System.Threading.AbandonedMutexException] {
    $mutexOk = $true
} catch {
    $mutexOk = $true
}
if (-not $mutexOk -and -not $Test -and -not $Preview) {
    try { Add-Content -Path (Join-Path $PSScriptRoot 'reminder.log') -Value ('[' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] 已有实例在运行，本次退出') -Encoding UTF8 } catch { }
    exit 0
}

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$selfPath = $PSCommandPath
$stateDir = Join-Path $root 'state'
$logFile = Join-Path $root 'reminder.log'
$libDates = Join-Path $root 'lib-dates.ps1'
if (Test-Path $libDates) { . $libDates }
$S = Get-Content (Join-Path $root 'strings.json') -Raw -Encoding UTF8 | ConvertFrom-Json

function Write-Log([string]$msg) {
    try {
        Add-Content -Path $script:logFile -Value ("[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg) -Encoding UTF8
    } catch { }
}

# Push the done/skip marks from state\*.txt into the workbook K column,
# so the user never has to tick the sheet by hand.
# The original file is never edited in place: work on a temp copy, then replace.
function Sync-Workbook {
    param([string]$Root, [string]$StateDir)
    $xlsx = Get-ChildItem -Path $Root -Filter '*.xlsx' -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $xlsx) { return 'no-workbook' }
    if (-not (Test-Path $StateDir)) { return 'no-state' }
    try {
        $planData = Get-Content (Join-Path $Root 'words.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch { return 'no-plan' }

    $vals = @{}
    $i = 0
    foreach ($d in $planData.days) {
        $i++
        $f = Join-Path $StateDir ($d.date + '.txt')
        if (Test-Path $f) {
            $v = (Get-Content $f -Raw).Trim()
            if ($v -eq 'done') { $vals[$i] = 1 }
            elseif ($v -eq 'skip') { $vals[$i] = 2 }
        }
    }
    if ($vals.Count -eq 0) { return 'nothing-to-sync' }

    Add-Type -AssemblyName System.IO.Compression.FileSystem, System.IO.Compression -ErrorAction SilentlyContinue
    $tmp = Join-Path $env:TEMP ('cet6_sync_' + [guid]::NewGuid().ToString('N') + '.xlsx')
    try {
        Copy-Item $xlsx.FullName $tmp -Force
        $zip = [System.IO.Compression.ZipFile]::Open($tmp, [System.IO.Compression.ZipArchiveMode]::Update)
        try {
            $entry = $zip.GetEntry('xl/worksheets/sheet1.xml')
            if (-not $entry) { return 'no-sheet' }
            $sr = New-Object System.IO.StreamReader($entry.Open())
            $xml = $sr.ReadToEnd()
            $sr.Close()
            $entry.Delete()
            foreach ($k in @($vals.Keys)) {
                $row = [int]$k + 1
                $ch = [string][char]0x2714
                if ($vals[$k] -eq 2) { $ch = [string][char]0x23ED }
                $styleAttr = ''
                $sm = [regex]::Match($xml, '<c r="H' + $row + '"[^>]*? s="(\d+)"')
                if ($sm.Success) { $styleAttr = ' s="' + $sm.Groups[1].Value + '"' }
                $cell = '<c r="H' + $row + '"' + $styleAttr + ' t="inlineStr"><is><t>' + $ch + '</t></is></c>'
                $pat = '<c r="H' + $row + '"[^>]*?(?:/>|>.*?</c>)'
                $xml = ([regex]::new($pat)).Replace($xml, $cell, 1)
            }
            $newEntry = $zip.CreateEntry('xl/worksheets/sheet1.xml')
            $sw = New-Object System.IO.StreamWriter($newEntry.Open(), (New-Object System.Text.UTF8Encoding($false)))
            $sw.Write($xml)
            $sw.Close()
        } finally { $zip.Dispose() }
        Copy-Item $tmp $xlsx.FullName -Force
        return 'ok'
    } catch {
        return ('fail: ' + $_.Exception.Message)
    } finally {
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    }
}

function Clear-WorkbookDay {
    param([string]$Root, [int]$Day)
    $xlsx = Get-ChildItem -Path $Root -Filter '*.xlsx' -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $xlsx) { return 'no-workbook' }
    $row = $Day + 1
    Add-Type -AssemblyName System.IO.Compression.FileSystem, System.IO.Compression -ErrorAction SilentlyContinue
    $tmp = Join-Path $env:TEMP ('cet6_clr_' + [guid]::NewGuid().ToString('N') + '.xlsx')
    try {
        Copy-Item $xlsx.FullName $tmp -Force
        $zip = [System.IO.Compression.ZipFile]::Open($tmp, [System.IO.Compression.ZipArchiveMode]::Update)
        try {
            $entry = $zip.GetEntry('xl/worksheets/sheet1.xml')
            if (-not $entry) { return 'no-sheet' }
            $sr = New-Object System.IO.StreamReader($entry.Open())
            $xml = $sr.ReadToEnd()
            $sr.Close()
            $entry.Delete()
            $styleAttr = ''
            $sm = [regex]::Match($xml, '<c r="H' + $row + '"[^>]*? s="(\d+)"')
            if ($sm.Success) { $styleAttr = ' s="' + $sm.Groups[1].Value + '"' }
            $cell = '<c r="H' + $row + '"' + $styleAttr + ' t="inlineStr"><is><t>' + [string][char]0x2610 + '</t></is></c>'
            $pat = '<c r="H' + $row + '"[^>]*?(?:/>|>.*?</c>)'
            $xml = ([regex]::new($pat)).Replace($xml, $cell, 1)
            $newEntry = $zip.CreateEntry('xl/worksheets/sheet1.xml')
            $sw = New-Object System.IO.StreamWriter($newEntry.Open(), (New-Object System.Text.UTF8Encoding($false)))
            $sw.Write($xml)
            $sw.Close()
        } finally { $zip.Dispose() }
        Copy-Item $tmp $xlsx.FullName -Force
        return 'ok'
    } catch {
        return ('fail: ' + $_.Exception.Message)
    } finally {
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    }
}

function Get-SubText {
    $text = ($script:S.subBase -f $script:today.ToString($script:S.dateFmt), $script:weekName, $script:streak)
    if ($script:targetOn) {
        $d = [int](($script:targetDate - $script:today).TotalDays)
        if ($d -ge 0) {
            $text += ($script:S.subTarget -f $script:targetLabel, $d)
        } else {
            $text += ('   已过' + $script:targetLabel + ' ' + (-$d) + ' 天')
        }
    }
    return $text
}

function Set-SubMode([string]$mode) {
    $script:subMode = $mode
    $pill = $script:win.FindName('SubPill')
    if ($mode -eq 'set') {
        $pill.Background = [Windows.Media.BrushConverter]::new().ConvertFromString('#1FFFFFFF')
        $pill.BorderBrush = [Windows.Media.BrushConverter]::new().ConvertFromString('#40FFFFFF')
    } else {
        $pill.Background = [Windows.Media.BrushConverter]::new().ConvertFromString('#33FFD98A')
        $pill.BorderBrush = [Windows.Media.BrushConverter]::new().ConvertFromString('#FFD98A')
    }
}

function Show-UndoBanner([string]$text) {
    $sub = $script:win.FindName('SubText')
    $sub.Text = $text
    $sub.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#FFD98A')
    Set-SubMode 'undo'
}

function Show-PausedBanner([string]$text) {
    $sub = $script:win.FindName('SubText')
    $sub.Text = $text
    $sub.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#FFD98A')
    Set-SubMode 'pause'
}

function Set-Pause {
    param([int]$Days)
    if (-not (Test-Path $script:stateDir)) { New-Item -ItemType Directory -Path $script:stateDir -Force | Out-Null }
    $f = Join-Path $script:stateDir 'pause.txt'
    $fromStr = $script:today.ToString('yyyy-MM-dd')
    if ($Days -le 0) { Set-Content $f ('forever|' + $fromStr) -Encoding UTF8 }
    else { Set-Content $f ($script:today.AddDays($Days).ToString('yyyy-MM-dd') + '|' + $fromStr) -Encoding UTF8 }
    Write-Log ('暂停提醒 ' + $Days + ' 天，计划一并冻结')
    $script:pauseLayer.Visibility = 'Collapsed'
    if ($Days -le 0) {
        Show-PausedBanner $script:S.pausedForever
    } else {
        Show-PausedBanner ($script:S.pausedBanner -f $script:today.AddDays($Days).ToString($script:S.dateFmt))
    }
    $script:statusText.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#2E7D4F')
    $script:statusText.Text = $script:S.pausedDone
}

function Resume-Pause {
    $f = Join-Path $script:stateDir 'pause.txt'
    $shift = 0
    if (Test-Path $f) {
        $raw = (Get-Content $f -Raw).Trim()
        $fromStr = $script:today.ToString('yyyy-MM-dd')
        if ($raw -match '^(.*?)\|(.*)$') { $fromStr = $matches[2] }
        try {
            $fromDate = [datetime]::ParseExact($fromStr, 'yyyy-MM-dd', $null)
            $shift = [int]($script:today - $fromDate).TotalDays
        } catch { $shift = 0 }
        if ($shift -lt 0) { $shift = 0 }
        Remove-Item $f -Force
    }
    if ($shift -gt 0) {
        try {
            [void](Fix-PlanDates -Dir $script:root -Start $script:planStart.AddDays($shift) -Exam $script:planExam)
            Write-Log ('手动恢复，计划顺延 ' + $shift + ' 天')
            $script:planStart = $script:planStart.AddDays($shift)
        } catch {
            Write-Log ('手动恢复顺延失败: ' + $_.Exception.Message)
        }
    }
    $sub = $script:win.FindName('SubText')
    $sub.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#CFE2F3')
    Set-SubMode 'set'
    if ($shift -gt 0) {
        $script:win.Close()
    } else {
        $sub.Text = (Get-SubText)
        $script:statusText.Text = $script:S.resumed
    }
}

function Undo-Today {
    $f = Join-Path $script:stateDir ($script:today.ToString('yyyy-MM-dd') + '.txt')
    if (Test-Path $f) { Remove-Item $f -Force }
    $r = Clear-WorkbookDay -Root $script:root -Day $script:dayIdx
    Write-Log ('undo today, workbook=' + $r)
    $script:streak = @(Get-ChildItem $script:stateDir -Filter '*.txt' | Where-Object {
                (Get-Content $_.FullName -Raw).Trim() -eq 'done'
            }).Count
    $sub = $script:win.FindName('SubText')
    $sub.Text = (Get-SubText)
    $sub.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#CFE2F3')
    Set-SubMode 'set'
    $script:statusText.Text = $script:S.undone
}

function Confirm-Saved([string]$text) {
    $script:statusText.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#2E7D4F')
    $script:statusText.Text = $text
    foreach ($c in $script:btnPanel.Children) { $c.IsEnabled = $false }

    # 打卡完成后给一个「继续背下一天」的入口，8 秒后自动关
    $nextBtn = New-Object Windows.Controls.Button
    $nextBtn.Content = $script:S.btnNextDay
    $nextBtn.Style = $script:win.Resources['FlatBtn']
    $nextBtn.Background = [Windows.Media.BrushConverter]::new().ConvertFromString('#EDF2F8')
    $nextBtn.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#1F4E79')
    $nextBtn.IsEnabled = $true
    $nextBtn.Add_Click({ param($s2, $e2) Start-NextDay })
    $script:btnPanel.Children.Add($nextBtn) | Out-Null

    $script:closeTimer = New-Object Windows.Threading.DispatcherTimer
    $script:closeTimer.Interval = [TimeSpan]::FromMilliseconds(8000)
    $script:closeTimer.Add_Tick({ $script:closeTimer.Stop(); $script:win.Close() })
    $script:closeTimer.Start()
}

if ($Sync) {
    $r = Sync-Workbook -Root $root -StateDir $stateDir
    Write-Log ('sync: ' + $r)
    Write-Output $r
    exit 0
}

try {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
    [System.Windows.Media.RenderOptions]::ProcessRenderMode = [System.Windows.Interop.RenderMode]::SoftwareOnly

    $plan = Get-Content (Join-Path $root 'words.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($AsOf) { $today = [datetime]::ParseExact($AsOf, 'yyyy-MM-dd', $null) } else { $today = (Get-Date).Date }
    $start = [datetime]::ParseExact($plan.start, 'yyyy-MM-dd', $null)
    $exam = [datetime]::ParseExact($plan.exam, 'yyyy-MM-dd', $null)
    $totalDays = @($plan.days).Count
    $dayIdx = [int]($today - $start).Days + 1

    $stateToday = ''
    $stampFile = Join-Path $stateDir ($today.ToString('yyyy-MM-dd') + '.txt')
    if (Test-Path $stampFile) { $stateToday = (Get-Content $stampFile -Raw).Trim() }

    $targetOn = $false
    $targetDate = $today
    $targetLabel = $S.targetDefault
    $script:uiMode = 'normal'
    $script:uiModeSaved = 'normal'
    $cfgFile = Join-Path $root 'config.json'
    if (Test-Path $cfgFile) {
        try {
            $cfg = Get-Content $cfgFile -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($cfg.target -and $cfg.target -ne 'none') {
                $targetDate = [datetime]::ParseExact($cfg.target, 'yyyy-MM-dd', $null)
                $targetOn = $true
            }
            if ($cfg.label) { $targetLabel = $cfg.label }
            if ($cfg.ui -eq 'compact') { $script:uiModeSaved = 'compact' }
        } catch { }
    }

    $pausedUntil = ''
    $pausedFrom = ''
    $pauseFile = Join-Path $stateDir 'pause.txt'
    if (Test-Path $pauseFile) {
        $rawPause = (Get-Content $pauseFile -Raw).Trim()
        if ($rawPause -match '^(.*?)\|(.*)$') {
            $pausedUntil = $matches[1]
            $pausedFrom = $matches[2]
        } else {
            $pausedUntil = $rawPause
            $pausedFrom = $today.ToString('yyyy-MM-dd')
        }
    }

    $isPaused = $false
    if ($pausedUntil) {
        try {
            $fromDate = [datetime]::ParseExact($pausedFrom, 'yyyy-MM-dd', $null)
        } catch {
            $fromDate = $today
        }
        if ($pausedUntil -eq 'forever') {
            $isPaused = $true
        } else {
            $untilDate = [datetime]::ParseExact($pausedUntil, 'yyyy-MM-dd', $null)
            if ($today -ge $untilDate) {
                # 暂停到期：计划整体顺延，然后清掉暂停标记
                $shift = [int]($untilDate - $fromDate).TotalDays
                if ($shift -gt 0) {
                    try {
                        [void](Fix-PlanDates -Dir $root -Start $start.AddDays($shift) -Exam $exam)
                        Write-Log ('暂停到期，计划顺延 ' + $shift + ' 天')
                        $plan = Get-Content (Join-Path $root 'words.json') -Raw -Encoding UTF8 | ConvertFrom-Json
                        $start = [datetime]::ParseExact($plan.start, 'yyyy-MM-dd', $null)
                        $exam = [datetime]::ParseExact($plan.exam, 'yyyy-MM-dd', $null)
                        $dayIdx = [int]($today - $start).Days + 1
                        $daysLeft = [int](($exam - $today).TotalDays)
                    } catch {
                        Write-Log ('顺延失败: ' + $_.Exception.Message)
                    }
                }
                Remove-Item $pauseFile -Force -ErrorAction SilentlyContinue
                $pausedUntil = ''
            } else {
                $isPaused = $true
            }
        }
    }

    # catch-up: last time the workbook may have been locked by Excel
    $syncRes = Sync-Workbook -Root $root -StateDir $stateDir
    if ($syncRes -notin @('ok', 'nothing-to-sync', 'no-state')) { Write-Log ('sync: ' + $syncRes) }

    if (-not $Preview -and -not $Test) {
        if ($dayIdx -lt 1) { exit 0 }
        if ($isPaused) { Write-Log ('暂停中（到 ' + $pausedUntil + '），不弹窗'); exit 0 }
        $stampFile = Join-Path $stateDir ($today.ToString('yyyy-MM-dd') + '.txt')
        if (Test-Path $stampFile) {
            $st = (Get-Content $stampFile -Raw).Trim()
            if ($st -eq 'done' -or $st -eq 'skip') { Write-Log $S.logAlready; exit 0 }
        }
    }

    if (-not (Test-Path $stateDir)) { New-Item -ItemType Directory -Path $stateDir -Force | Out-Null }

    # ---------- data ----------
    $isFinished = $dayIdx -gt $totalDays
    $day = $null
    if (-not $isFinished) { $day = $plan.days[$dayIdx - 1] }

    $groups = @()
    if ($null -ne $day) {
        foreach ($iv in $plan.intervals) {
            $src = $dayIdx - [int]$iv
            if ($src -ge 1) {
                $groups += [pscustomobject]@{
                    Label = ($S.groupLabel -f $iv, $src)
                    Words = @($plan.days[$src - 1].words)
                }
            }
        }
    }

    $streak = 0
    if (Test-Path $stateDir) {
        $streak = @(Get-ChildItem $stateDir -Filter '*.txt' | Where-Object {
                    (Get-Content $_.FullName -Raw).Trim() -eq 'done'
                }).Count
    }
    $daysLeft = [int](($exam - $today).TotalDays)
    if ($daysLeft -lt 0 -and -not $Preview -and -not $Test) { exit 0 }
    $weekCn = @(
        (-join ([char]0x5468, [char]0x65E5)),
        (-join ([char]0x5468, [char]0x4E00)),
        (-join ([char]0x5468, [char]0x4E8C)),
        (-join ([char]0x5468, [char]0x4E09)),
        (-join ([char]0x5468, [char]0x56DB)),
        (-join ([char]0x5468, [char]0x4E94)),
        (-join ([char]0x5468, [char]0x516D))
    )
    $weekName = $weekCn[[int]$today.DayOfWeek]

    # ---------- window ----------
    $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        ResizeMode="NoResize" SizeToContent="Height" Width="900" Title="CET-6 Word Reminder"
        WindowStartupLocation="CenterScreen" ShowInTaskbar="True" Topmost="True"
        FontFamily="Microsoft YaHei UI">
  <Window.Resources>
    <Style x:Key="FlatBtn" TargetType="Button">
      <Setter Property="Padding" Value="14,8"/>
      <Setter Property="Margin" Value="7,0,0,0"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" CornerRadius="8" Background="{TemplateBinding Background}"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Opacity" Value="0.82"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>
  <Border CornerRadius="14" Background="#FFFFFF" Margin="0">
    <Border.LayoutTransform>
      <ScaleTransform ScaleX="1.5" ScaleY="1.5"/>
    </Border.LayoutTransform>
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>
      <Border Grid.Row="0" CornerRadius="16,16,0,0" Name="Header">
        <Border.Background>
          <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
            <GradientStop Color="#1F4E79" Offset="0"/>
            <GradientStop Color="#2E75B6" Offset="1"/>
          </LinearGradientBrush>
        </Border.Background>
        <Grid Name="HeaderGrid" Margin="26,18,18,18">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <StackPanel Grid.Column="0">
            <TextBlock Name="TitleText" Text="" Foreground="#FFFFFF" FontSize="19" FontWeight="Bold"/>
            <StackPanel Orientation="Horizontal" Margin="0,6,0,0" HorizontalAlignment="Left">
              <Border Name="SubPill" CornerRadius="11" Padding="10,3" Background="#1FFFFFFF"
                      BorderBrush="#40FFFFFF" BorderThickness="1" Cursor="Hand">
                <TextBlock Name="SubText" Text="" Foreground="#CFE2F3" FontSize="12"/>
              </Border>
              <Border Name="GearBtn" Width="22" Height="22" CornerRadius="11" Margin="8,0,0,0"
                      Background="#26FFFFFF" Cursor="Hand" ToolTip="设置：目标日期 / 每天词数">
                <TextBlock Text="&#x2699;" FontFamily="Segoe UI Symbol" FontSize="12"
                           Foreground="#FFFFFF" HorizontalAlignment="Center" VerticalAlignment="Center"/>
              </Border>
            </StackPanel>
          </StackPanel>
          <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Top">
            <Border Width="30" Height="30" CornerRadius="15" Background="#33FFFFFF"
                    Margin="0,0,8,0" Cursor="Hand" Name="NarrowBtn"
                    ToolTip="窄条模式：一行一个词，不挡其他窗口">
              <TextBlock Text="窄" Foreground="#FFFFFF" FontSize="11" FontWeight="Bold"
                         HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <Border Width="30" Height="30" CornerRadius="15" Background="#33FFFFFF"
                    Margin="0,0,8,0" Cursor="Hand" Name="DeckBtn" ToolTip="切换词库">
              <TextBlock Text="词" Foreground="#FFFFFF" FontSize="11" FontWeight="Bold"
                         HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <Border Width="30" Height="30" CornerRadius="15" Background="#33FFFFFF"
                    Margin="0,0,8,0" Cursor="Hand" Name="PauseBtn" ToolTip="暂停自动提醒">
              <TextBlock Text="II" Foreground="#FFFFFF" FontSize="12" FontWeight="Bold"
                         HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <Border Width="30" Height="30" CornerRadius="15" Background="#33FFFFFF"
                    Margin="0,0,8,0" Cursor="Hand" Name="WideBtn" Visibility="Collapsed"
                    ToolTip="回到标准模式">
              <TextBlock Text="宽" Foreground="#FFFFFF" FontSize="11" FontWeight="Bold"
                         HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <Border Width="30" Height="30" CornerRadius="15" Background="#33FFFFFF"
                    Margin="0,0,8,0" Cursor="Hand" Name="MinBtn"
                    ToolTip="最小化（收进任务栏，点任务栏图标可恢复）">
              <Rectangle Width="12" Height="2" RadiusX="1" RadiusY="1" Fill="#FFFFFF"
                         HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <Border Width="30" Height="30" CornerRadius="15" Background="#33FFFFFF"
                    VerticalAlignment="Top" Cursor="Hand" Name="CloseBtn">
              <TextBlock Text="&#x2715;" Foreground="#FFFFFF" FontSize="13" HorizontalAlignment="Center"
                         VerticalAlignment="Center"/>
            </Border>
          </StackPanel>
        </Grid>
      </Border>
      <ScrollViewer Name="MainScroll" Grid.Row="1" VerticalScrollBarVisibility="Auto" MaxHeight="430">
        <StackPanel Name="BodyPanel" Margin="26,16,26,6"/>
      </ScrollViewer>
      <ScrollViewer Name="CompactScroll" Grid.Row="1" VerticalScrollBarVisibility="Auto" MaxHeight="430"
                    Visibility="Collapsed">
        <StackPanel Name="CompactPanel" Margin="16,14,16,6"/>
      </ScrollViewer>
      <Grid Grid.Row="2" Margin="26,8,26,22">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock Name="StatusText" Grid.Column="0" VerticalAlignment="Center" Foreground="#2E7D4F"
                   FontSize="12" FontWeight="Bold" Text="" TextWrapping="Wrap"/>
        <StackPanel Grid.Column="1" Orientation="Horizontal" Name="BtnPanel"/>
      </Grid>
      <Grid Name="CompactFoot" Grid.Row="2" Margin="16,6,16,14" Visibility="Collapsed">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock Name="CompactStatus" Grid.Column="0" VerticalAlignment="Center" Foreground="#2E7D4F"
                   FontSize="11" FontWeight="Bold" Text="" TextWrapping="Wrap"/>
        <StackPanel Grid.Column="1" Orientation="Horizontal" Name="CompactBtnPanel"/>
      </Grid>
      <Border Name="SetLayer" Grid.Row="0" Grid.RowSpan="3" CornerRadius="16"
              Background="#8C1B2A3A" Visibility="Collapsed">
        <Border Background="#FFFFFF" CornerRadius="12" Padding="26,22" MaxWidth="480"
                VerticalAlignment="Center" HorizontalAlignment="Center">
          <StackPanel>
            <TextBlock Name="SetTitle" Text="" FontSize="17" FontWeight="Bold" Foreground="#1A1A1A"/>
            <TextBlock Name="SetDateHint" Text="" FontSize="11.5" Foreground="#999999" Margin="0,10,0,6"/>
            <TextBox Name="SetDate" FontSize="14" Padding="8,6" FontFamily="Consolas"/>
            <TextBlock Name="SetLabelHint" Text="" FontSize="11.5" Foreground="#999999" Margin="0,12,0,6"/>
            <TextBox Name="SetLabel" FontSize="14" Padding="8,6"/>
            <TextBlock Name="SetPerDayHint" Text="" FontSize="11.5" Foreground="#999999" Margin="0,12,0,6"/>
            <TextBox Name="SetPerDay" FontSize="14" Padding="8,6" FontFamily="Consolas" Width="120" HorizontalAlignment="Left"/>
            <TextBlock Name="SetFeedback" Text="" FontSize="11.5" Foreground="#C0392B" Margin="0,10,0,0"/>
            <Button Name="SetReset" Content="" Style="{StaticResource FlatBtn}" HorizontalAlignment="Left"
                    Background="#FDECEA" Foreground="#C0392B" Margin="0,12,0,0"/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0">
              <Button Name="SetSave" Content="" Style="{StaticResource FlatBtn}" Background="#1F4E79" Foreground="#FFFFFF"/>
              <Button Name="SetCancel" Content="" Style="{StaticResource FlatBtn}" Background="#FFFFFF" Foreground="#999999"/>
            </StackPanel>
          </StackPanel>
        </Border>
      </Border>
      <Border Name="ResetLayer" Grid.Row="0" Grid.RowSpan="3" CornerRadius="16"
              Background="#8C1B2A3A" Visibility="Collapsed">
        <Border Background="#FFFFFF" CornerRadius="12" Padding="26,22" MaxWidth="460"
                VerticalAlignment="Center" HorizontalAlignment="Center">
          <StackPanel>
            <TextBlock Name="ResetTitle" Text="" FontSize="17" FontWeight="Bold" Foreground="#1A1A1A"/>
            <TextBlock Name="ResetBody" Text="" FontSize="12" Foreground="#666666"
                       TextWrapping="Wrap" LineHeight="19" Margin="0,10,0,0"/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
              <Button Name="ResetNo" Content="" Style="{StaticResource FlatBtn}"
                      Background="#EDF2F8" Foreground="#1F4E79"/>
              <Button Name="ResetYes" Content="" Style="{StaticResource FlatBtn}"
                      Background="#C0392B" Foreground="#FFFFFF"/>
            </StackPanel>
          </StackPanel>
        </Border>
      </Border>
      <Border Name="DeckLayer" Grid.Row="0" Grid.RowSpan="3" CornerRadius="16"
              Background="#8C1B2A3A" Visibility="Collapsed">
        <Border Background="#FFFFFF" CornerRadius="12" Padding="26,22" MaxWidth="520"
                VerticalAlignment="Center" HorizontalAlignment="Center">
          <StackPanel>
            <TextBlock Name="DeckTitle" Text="" FontSize="17" FontWeight="Bold" Foreground="#1A1A1A"/>
            <TextBlock Name="DeckHint" Text="" FontSize="11.5" Foreground="#999999"
                       TextWrapping="Wrap" Margin="0,8,0,12"/>
            <StackPanel Name="DeckList"/>
            <TextBlock Name="DeckFormat" Text="" FontSize="10.5" Foreground="#8C8C8C"
                       TextWrapping="Wrap" Margin="0,12,0,0" LineHeight="16" Visibility="Collapsed"/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
              <Button Name="DeckHelp" Content="" Style="{StaticResource FlatBtn}"
                      Background="Transparent" Foreground="#8C8C8C" FontSize="11" Padding="8,6"/>
              <Button Name="DeckAdd" Content="" Style="{StaticResource FlatBtn}"
                      Background="#EDF2F8" Foreground="#1F4E79"/>
              <Button Name="DeckClose" Content="" Style="{StaticResource FlatBtn}"
                      Background="#FFFFFF" Foreground="#999999"/>
            </StackPanel>
          </StackPanel>
        </Border>
      </Border>
      <Border Name="PauseLayer" Grid.Row="0" Grid.RowSpan="3" CornerRadius="16"
              Background="#8C1B2A3A" Visibility="Collapsed">
        <Border Background="#FFFFFF" CornerRadius="12" Padding="26,22" MaxWidth="460"
                VerticalAlignment="Center" HorizontalAlignment="Center">
          <StackPanel>
            <TextBlock Name="PauseTitle" Text="" FontSize="17" FontWeight="Bold" Foreground="#1A1A1A"/>
            <TextBlock Name="PauseBody" Text="" FontSize="12" Foreground="#666666"
                       TextWrapping="Wrap" Margin="0,10,0,16"/>
            <WrapPanel>
              <Button Name="PauseD1" Content="" Style="{StaticResource FlatBtn}" Background="#1F4E79" Foreground="#FFFFFF"/>
              <Button Name="PauseD3" Content="" Style="{StaticResource FlatBtn}" Background="#1F4E79" Foreground="#FFFFFF"/>
              <Button Name="PauseD7" Content="" Style="{StaticResource FlatBtn}" Background="#1F4E79" Foreground="#FFFFFF"/>
              <Button Name="PauseAlways" Content="" Style="{StaticResource FlatBtn}" Background="#EDF2F8" Foreground="#1F4E79"/>
            </WrapPanel>
            <Button Name="PauseCancel" Content="" Style="{StaticResource FlatBtn}" Background="#FFFFFF"
                    Foreground="#999999" HorizontalAlignment="Right" Margin="0,16,0,0"/>
          </StackPanel>
        </Border>
      </Border>
      <Border Name="ConfirmLayer" Grid.Row="0" Grid.RowSpan="3" CornerRadius="16"
              Background="#8C1B2A3A" Visibility="Collapsed">
        <Border Background="#FFFFFF" CornerRadius="12" Padding="26,22" MaxWidth="430"
                VerticalAlignment="Center" HorizontalAlignment="Center">
          <StackPanel>
            <TextBlock Name="ConfirmTitle" Text="" FontSize="17" FontWeight="Bold" Foreground="#1A1A1A"/>
            <TextBlock Name="ConfirmBody" Text="" FontSize="12" Foreground="#666666"
                       TextWrapping="Wrap" Margin="0,10,0,18"/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
              <Button Name="BtnSkipNo" Content="" Style="{StaticResource FlatBtn}"
                      Background="#EDF2F8" Foreground="#1F4E79"/>
              <Button Name="BtnSkipYes" Content="" Style="{StaticResource FlatBtn}"
                      Background="#C0392B" Foreground="#FFFFFF"/>
            </StackPanel>
          </StackPanel>
        </Border>
      </Border>
    </Grid>
  </Border>
</Window>
'@

    $win = [Windows.Markup.XamlReader]::Parse($xaml)
    $body = $win.FindName('BodyPanel')
    $btnPanel = $win.FindName('BtnPanel')
    $statusText = $win.FindName('StatusText')
    $script:pauseHandlerSet = $false
    $script:undoHandlerSet = $false
    $confirmLayer = $win.FindName('ConfirmLayer')
    $setLayer = $win.FindName('SetLayer')
    $resetLayer = $win.FindName('ResetLayer')
    $win.FindName('ResetTitle').Text = $S.resetTitle
    $win.FindName('ResetBody').Text = $S.resetBody
    $win.FindName('ResetYes').Content = $S.resetYes
    $win.FindName('ResetNo').Content = $S.resetNo
    $win.FindName('SetReset').Content = $S.resetBtn
    $win.FindName('SetTitle').Text = $S.setTitle
    $win.FindName('SetDateHint').Text = $S.setDateHint
    $win.FindName('SetLabelHint').Text = $S.setLabelHint
    $win.FindName('SetPerDayHint').Text = $S.setPerDayHint
    $win.FindName('SetSave').Content = $S.setSave
    $win.FindName('SetCancel').Content = $S.setCancel

    $deckLayer = $win.FindName('DeckLayer')
    $deckList = $win.FindName('DeckList')
    $win.FindName('DeckTitle').Text = $S.deckTitle
    $win.FindName('DeckHint').Text = $S.deckHint
    $win.FindName('DeckAdd').Content = $S.deckAdd
    $win.FindName('DeckClose').Content = $S.deckClose
    $win.FindName('DeckFormat').Text = $S.deckFormat
    $win.FindName('DeckHelp').Content = $S.deckHelp

    $pauseLayer = $win.FindName('PauseLayer')
    $win.FindName('PauseTitle').Text = $S.pauseTitle
    $win.FindName('PauseBody').Text = $S.pauseBody
    $win.FindName('PauseD1').Content = $S.pauseD1
    $win.FindName('PauseD3').Content = $S.pauseD3
    $win.FindName('PauseD7').Content = $S.pauseD7
    $win.FindName('PauseAlways').Content = $S.pauseAlways
    $win.FindName('PauseCancel').Content = $S.pauseCancel
    $win.FindName('ConfirmTitle').Text = $S.confirmSkipTitle
    $win.FindName('ConfirmBody').Text = $S.confirmSkipBody
    $win.FindName('BtnSkipNo').Content = $S.confirmSkipNo
    $win.FindName('BtnSkipYes').Content = $S.confirmSkipYes

    $win.FindName('Header').Add_MouseLeftButtonDown({ try { $script:win.DragMove() } catch { } })
    $win.FindName('NarrowBtn').Add_MouseLeftButtonDown({ param($s2, $e2) $e2.Handled = $true; Set-UiMode 'compact' })
    $win.FindName('WideBtn').Add_MouseLeftButtonDown({ param($s2, $e2) $e2.Handled = $true; Set-UiMode 'normal' })
    # the close button sits inside the drag area: swallow the press so DragMove()
    # does not eat the event (that is why clicking X did nothing before)
    $minBtn = $win.FindName('MinBtn')
    $minBtn.Add_MouseLeftButtonDown({ param($s, $e) $e.Handled = $true; $script:win.WindowState = 'Minimized' })
    $minBtn.Add_MouseLeftButtonUp({ param($s, $e) $e.Handled = $true })
    $closeBtn = $win.FindName('CloseBtn')
    $closeBtn.Add_MouseLeftButtonDown({ param($s, $e) $e.Handled = $true; $script:win.Close() })
    $closeBtn.Add_MouseLeftButtonUp({ param($s, $e) $e.Handled = $true })

    function New-Text([string]$text, [double]$size, [string]$color, [string]$weight = 'Normal', [double]$marginTop = 0) {
        $t = New-Object Windows.Controls.TextBlock
        $t.Text = $text
        $t.FontSize = $size
        $t.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString($color)
        $t.FontWeight = [Windows.FontWeights]::$weight
        $t.TextWrapping = 'Wrap'
        $t.Margin = New-Object Windows.Thickness(0, $marginTop, 0, 0)
        return $t
    }

    function New-Sel([string]$text, [double]$size, [string]$color, [string]$weight = 'Normal', [double]$marginTop = 0, [string]$align = 'Stretch') {
        $t = New-Object Windows.Controls.TextBox
        $t.Text = $text
        $t.FontSize = $size
        $t.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString($color)
        $t.FontWeight = [Windows.FontWeights]::$weight
        $t.TextWrapping = 'Wrap'
        $t.IsReadOnly = $true
        $t.IsTabStop = $false
        $t.BorderThickness = New-Object Windows.Thickness(0)
        $t.Background = [Windows.Media.Brushes]::Transparent
        $t.Padding = New-Object Windows.Thickness(0)
        $t.Margin = New-Object Windows.Thickness(0, $marginTop, 0, 0)
        $t.Cursor = [Windows.Input.Cursors]::IBeam
        $t.VerticalContentAlignment = 'Center'
        if ($align -eq 'Left') { $t.HorizontalAlignment = 'Left' }
        return $t
    }

    function Add-CopyMenu($element, [string]$value) {
        $cm = New-Object Windows.Controls.ContextMenu
        $mi = New-Object Windows.Controls.MenuItem
        $mi.Header = $script:S.copyOne
        $mi.Add_Click({ param($s2, $e2) [Windows.Clipboard]::SetText($value) }.GetNewClosure())
        $cm.Items.Add($mi) | Out-Null
        $element.ContextMenu = $cm
    }

    function Add-Btn([string]$text, [string]$bg, [string]$fg, [scriptblock]$action, $panel) {
        $b = New-Object Windows.Controls.Button
        $b.Content = $text
        $b.Style = $script:win.Resources['FlatBtn']
        $b.Background = [Windows.Media.BrushConverter]::new().ConvertFromString($bg)
        $b.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString($fg)
        $b.Add_Click($action)
        if (-not $panel) { $panel = $script:btnPanel }
        $panel.Children.Add($b) | Out-Null
        $script:lastBtn = $b
    }

    function New-Line([string]$text, [double]$size, [string]$color, [string]$weight = 'Normal', [double]$left = 0) {
        $t = New-Object Windows.Controls.TextBlock
        $t.Text = $text
        $t.FontSize = $size
        $t.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString($color)
        $t.FontWeight = [Windows.FontWeights]::$weight
        $t.TextWrapping = 'NoWrap'
        $t.TextTrimming = 'CharacterEllipsis'
        $t.Margin = New-Object Windows.Thickness($left, 0, 0, 0)
        $t.VerticalAlignment = 'Center'
        return $t
    }

    # ---------- content ----------
    if ($isFinished) {
        $script:titleFull = $S.titleDone
        $script:titleShort = $S.titleDone
        $win.FindName('TitleText').Text = $script:titleFull
        if ($daysLeft -ge 0) { $win.FindName('SubText').Text = ($S.subDone -f $daysLeft) } else { $win.FindName('SubText').Text = $S.subEnded }
        $body.Children.Add((New-Text $S.doneBody1 14 '#333333')) | Out-Null
        $body.Children.Add((New-Text $S.doneBody2 12 '#666666' 'Normal' 10)) | Out-Null
        $body.Children.Add((New-Text $S.doneBody3 12 '#666666' 'Normal' 6)) | Out-Null
    }
    else {
        $script:titleFull = ($S.titleToday -f $day.day)
        $script:titleShort = ('第 {0} 天' -f $day.day)
        $win.FindName('TitleText').Text = $script:titleFull
        $script:targetOn = $targetOn
        $script:targetDate = $targetDate
        $script:targetLabel = $targetLabel
        $script:planStart = $start
        $script:planPerDay = [int]$plan.perDay
        $script:planIntervals = @($plan.intervals)
        $script:planExam = $exam
        $script:subMode = 'set'
        $pill0 = $win.FindName('SubPill')
        $pill0.Add_MouseLeftButtonDown({
                param($s2, $e2)
                $e2.Handled = $true
                if ($script:subMode -eq 'undo') { Undo-Today }
                elseif ($script:subMode -eq 'pause') { Resume-Pause }
                else { Open-SetLayer }
            })
        $pill0.Add_MouseEnter({
                $script:win.FindName('SubPill').Background = [Windows.Media.BrushConverter]::new().ConvertFromString(
                    $(if ($script:subMode -eq 'set') { '#33FFFFFF' } else { '#4DFFD98A' }))
            })
        $pill0.Add_MouseLeave({ Set-SubMode $script:subMode })
        $gear0 = $win.FindName('GearBtn')
        $gear0.Add_MouseLeftButtonDown({
                param($s2, $e2)
                $e2.Handled = $true
                Open-SetLayer
            })
        $gear0.Add_MouseEnter({
                $script:win.FindName('GearBtn').Background = [Windows.Media.BrushConverter]::new().ConvertFromString('#4DFFFFFF')
            })
        $gear0.Add_MouseLeave({
                $script:win.FindName('GearBtn').Background = [Windows.Media.BrushConverter]::new().ConvertFromString('#26FFFFFF')
            })
        $win.FindName('SubText').Text = (Get-SubText)

        if ($isPaused) {
            if ($pausedUntil -eq 'forever') {
                Show-PausedBanner $S.pausedForever
            } else {
                $pd = [datetime]::ParseExact($pausedUntil, 'yyyy-MM-dd', $null)
                Show-PausedBanner ($S.pausedBanner -f $pd.ToString('M月d日'))
            }
        } elseif ($stateToday -eq 'skip' -or $stateToday -eq 'done') {
            if ($stateToday -eq 'skip') { Show-UndoBanner $S.undoSkip } else { Show-UndoBanner $S.undoDone }
        }

        $body.Children.Add((New-Text $S.secNew 12 '#1F4E79' 'Bold')) | Out-Null
        $colCount = 2
        $rowCount = [int][math]::Ceiling($day.words.Count / $colCount)
        $grid = New-Object Windows.Controls.Grid
        $grid.Margin = New-Object Windows.Thickness(0, 2, 0, 0)
        for ($ci = 0; $ci -lt $colCount; $ci++) {
            $cd = New-Object Windows.Controls.ColumnDefinition
            $cd.Width = [Windows.GridLength]::new(1, [Windows.GridUnitType]::Star)
            $grid.ColumnDefinitions.Add($cd) | Out-Null
        }
        for ($ri = 0; $ri -lt $rowCount; $ri++) {
            $rd = New-Object Windows.Controls.RowDefinition
            $rd.Height = [Windows.GridLength]::new(1, [Windows.GridUnitType]::Auto)
            $grid.RowDefinitions.Add($rd) | Out-Null
        }
        $i = 0
        foreach ($w in $day.words) {
            $ci = [int][math]::Floor($i / $rowCount)
            $ri = $i % $rowCount
            $cell = New-Object Windows.Controls.StackPanel
            $cell.Margin = New-Object Windows.Thickness(0, 6, 14, 6)
            $top = New-Object Windows.Controls.StackPanel
            $top.Orientation = 'Horizontal'
            $wTb = New-Sel $w.w 14 '#1A1A1A' 'Bold' 0 'Left'
            $wTb.VerticalAlignment = 'Center'
            $top.Children.Add($wTb) | Out-Null
            $pos = New-Sel $w.pos 9.5 '#999999' 'Normal' 0 'Left'
            $pos.VerticalAlignment = 'Center'
            $pos.Margin = New-Object Windows.Thickness(5, 2, 0, 0)
            $top.Children.Add($pos) | Out-Null
            $cell.Children.Add($top) | Out-Null
            if ($w.ipa) {
                $ipaTb = New-Sel $w.ipa 11 '#5B7A99' 'Normal' 1 'Left'
                $ipaTb.FontFamily = New-Object Windows.Media.FontFamily('Lucida Sans Unicode')
                $cell.Children.Add($ipaTb) | Out-Null
            }
            $cnTb = New-Sel $w.cn 12 '#333333' 'Normal' 2
            $cell.Children.Add($cnTb) | Out-Null
            if ($w.ph) {
                $phTb = New-Sel $w.ph 9.5 '#8C8C8C' 'Normal' 1
                $cell.Children.Add($phTb) | Out-Null
            }
            [Windows.Controls.Grid]::SetColumn($cell, $ci)
            [Windows.Controls.Grid]::SetRow($cell, $ri)
            $grid.Children.Add($cell) | Out-Null
            $i++
        }
        $body.Children.Add($grid) | Out-Null

        if ($groups.Count -gt 0) {
            $line = New-Object Windows.Controls.Border
            $line.Height = 1
            $line.Background = [Windows.Media.BrushConverter]::new().ConvertFromString('#EDEDED')
            $line.Margin = New-Object Windows.Thickness(0, 14, 0, 12)
            $body.Children.Add($line) | Out-Null

            $totalReview = 0
            foreach ($gp in $groups) { $totalReview += $gp.Words.Count }

            $revHead = New-Object Windows.Controls.Grid
            $rc1 = New-Object Windows.Controls.ColumnDefinition
            $rc1.Width = New-Object Windows.GridLength(1, [Windows.GridUnitType]::Star)
            $rc2 = New-Object Windows.Controls.ColumnDefinition
            $rc2.Width = [Windows.GridLength]::Auto
            $revHead.ColumnDefinitions.Add($rc1)
            $revHead.ColumnDefinitions.Add($rc2)

            $revTitle = New-Text ($S.secReview -f $groups.Count, $totalReview) 12 '#1F4E79' 'Bold'
            $revTitle.VerticalAlignment = 'Center'
            [Windows.Controls.Grid]::SetColumn($revTitle, 0)
            $revHead.Children.Add($revTitle) | Out-Null

            $script:revExpanded = $false
            $revToggle = New-Object Windows.Controls.Button
            $revToggle.Content = $script:S.revShow
            $revToggle.Style = $win.Resources['FlatBtn']
            $revToggle.FontSize = 10.5
            $revToggle.Padding = New-Object Windows.Thickness(10, 2, 10, 2)
            $revToggle.Background = [Windows.Media.BrushConverter]::new().ConvertFromString('#EDF2F8')
            $revToggle.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#1F4E79')
            $revToggle.VerticalAlignment = 'Center'
            $revToggle.HorizontalAlignment = 'Right'
            [Windows.Controls.Grid]::SetColumn($revToggle, 1)
            $revHead.Children.Add($revToggle) | Out-Null
            $body.Children.Add($revHead) | Out-Null

            $flatList = New-Object System.Collections.ArrayList
            $detailList = New-Object System.Collections.ArrayList

            foreach ($gp in $groups) {
                $body.Children.Add((New-Text $gp.Label 10.5 '#999999' 'Bold' 9)) | Out-Null

                $txt = ($gp.Words | ForEach-Object { $_.w }) -join $S.wordSep
                $flat = New-Sel $txt 11 '#555555' 'Normal' 3
                $body.Children.Add($flat) | Out-Null
                [void]$flatList.Add($flat)

                $det = New-Object Windows.Controls.StackPanel
                $det.Visibility = 'Collapsed'
                $det.Margin = New-Object Windows.Thickness(0, 2, 0, 4)
                foreach ($rw in $gp.Words) {
                    $p2 = New-Object Windows.Controls.TextBlock
                    $p2.TextWrapping = 'Wrap'
                    $p2.Margin = New-Object Windows.Thickness(0, 3, 0, 3)
                    $r1 = New-Object Windows.Documents.Run($rw.w)
                    $r1.FontWeight = 'Bold'
                    $r1.FontSize = 11.5
                    $r1.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#1A1A1A')
                    $p2.Inlines.Add($r1) | Out-Null
                    if ($rw.pos) {
                        $rp = New-Object Windows.Documents.Run(('  ' + $rw.pos))
                        $rp.FontSize = 9.5
                        $rp.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#999999')
                        $p2.Inlines.Add($rp) | Out-Null
                    }
                    if ($rw.ipa) {
                        $r2 = New-Object Windows.Documents.Run(('  ' + $rw.ipa))
                        $r2.FontFamily = New-Object Windows.Media.FontFamily('Lucida Sans Unicode')
                        $r2.FontSize = 10
                        $r2.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#5B7A99')
                        $p2.Inlines.Add($r2) | Out-Null
                    }
                    $r3 = New-Object Windows.Documents.Run(('   ' + $rw.cn))
                    $r3.FontSize = 11
                    $r3.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#555555')
                    $p2.Inlines.Add($r3) | Out-Null
                    $rowText = $rw.w + $(if ($rw.pos) { '  ' + $rw.pos } else { '' }) + $(if ($rw.ipa) { '  ' + $rw.ipa } else { '' }) + '  ' + $rw.cn
                    Add-CopyMenu $p2 $rowText
                    $det.Children.Add($p2) | Out-Null
                }
                $body.Children.Add($det) | Out-Null
                [void]$detailList.Add($det)
            }

            $revToggle.Add_Click({
                    if ($script:revExpanded) {
                        foreach ($t in $flatList) { $t.Visibility = 'Visible' }
                        foreach ($t in $detailList) { $t.Visibility = 'Collapsed' }
                        $script:revExpanded = $false
                        $revToggle.Content = $script:S.revShow
                    } else {
                        foreach ($t in $flatList) { $t.Visibility = 'Collapsed' }
                        foreach ($t in $detailList) { $t.Visibility = 'Visible' }
                        $script:revExpanded = $true
                        $revToggle.Content = $script:S.revHide
                    }
                }.GetNewClosure())
        }
    }

    function Open-SetLayer {
        $win.FindName('SetFeedback').Text = ''
        if ($script:targetOn) {
            $win.FindName('SetDate').Text = $script:targetDate.ToString('yyyy-MM-dd')
        } else {
            $win.FindName('SetDate').Text = ''
        }
        $win.FindName('SetLabel').Text = $script:targetLabel
        $win.FindName('SetPerDay').Text = [string]$script:planPerDay
        $script:setLayer.Visibility = 'Visible'
        $win.FindName('SetDate').Focus() | Out-Null
    }

    function Save-Target {
        $raw = $win.FindName('SetDate').Text.Trim()
        $label = $win.FindName('SetLabel').Text.Trim()
        if (-not $label) { $label = $script:S.targetDefault }

        $cfgFile = Join-Path $script:root 'config.json'
        $cfg = @{}
        if (Test-Path $cfgFile) {
            try {
                $old = Get-Content $cfgFile -Raw -Encoding UTF8 | ConvertFrom-Json
                foreach ($p in $old.PSObject.Properties) { $cfg[$p.Name] = $p.Value }
            } catch { }
        }

        if (-not $raw) {
            $cfg['target'] = 'none'
            $cfg['label'] = $label
            $script:targetOn = $false
        } else {
            try {
                $d2 = [datetime]::ParseExact($raw, 'yyyy-MM-dd', $null)
            } catch {
                $win.FindName('SetFeedback').Text = $script:S.setBad
                return
            }
            $cfg['target'] = $d2.ToString('yyyy-MM-dd')
            $cfg['label'] = $label
            $script:targetOn = $true
            $script:targetDate = $d2
        }
        $script:targetLabel = $label

        # 每天词数：改了才重建计划
        $perRaw = $win.FindName('SetPerDay').Text.Trim()
        $perNew = 0
        if ($perRaw) {
            if (-not [int]::TryParse($perRaw, [ref]$perNew) -or $perNew -lt 1 -or $perNew -gt 60) {
                $win.FindName('SetFeedback').Text = '每天词数要填 1–60 之间的整数'
                return
            }
        }

        ($cfg | ConvertTo-Json) | Set-Content $cfgFile -Encoding UTF8
        Write-Log ('目标日期: ' + $(if ($script:targetOn) { $cfg['target'] } else { '关闭' }))

        $sub = $script:win.FindName('SubText')
        $sub.Text = (Get-SubText)
        $script:statusText.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#2E7D4F')
        $script:statusText.Text = $script:S.setOk
        $script:setLayer.Visibility = 'Collapsed'

        if ($perNew -gt 0 -and $perNew -ne $script:planPerDay) {
            Invoke-RebuildPlan $perNew
        }
    }

    function Invoke-RebuildPlan([int]$perDay) {
        $py = Find-Python
        if (-not $py) {
            $script:statusText.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#C0392B')
            $script:statusText.Text = '找不到 Python，无法重排计划'
            return
        }
        $cfgFile = Join-Path $script:root 'config.json'
        $deckName = ''
        if (Test-Path $cfgFile) {
            try { $deckName = (Get-Content $cfgFile -Raw -Encoding UTF8 | ConvertFrom-Json).deck } catch { }
        }
        $deckPath = ''
        if ($deckName) { $deckPath = Join-Path $script:root ('词库\' + $deckName + '.csv') }
        if (-not (Test-Path $deckPath)) {
            $first = Get-ChildItem (Join-Path $script:root '词库') -Filter '*.csv' -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($first) { $deckPath = $first.FullName }
        }
        if (-not (Test-Path $deckPath)) {
            $script:statusText.Text = '找不到词库文件'
            return
        }
        $script:statusText.Text = '正在按新的每日词数重排…'
        $startStr = $script:planStart.ToString('yyyy-MM-dd')
        $commonArgs = @('--deck', $deckPath, '--per-day', $perDay, '--intervals', ($script:planIntervals -join ','),
            '--exam', $script:planExam.ToString('yyyy-MM-dd'))
        & $py (Join-Path $script:root '工具\build_plan.py') @commonArgs --start $startStr 2>&1 | Out-Null
        # 词数变多后计划可能已经走到今天之后，回夹让今天正好落在最后一天
        $np = Get-Content (Join-Path $script:root 'words.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        $lastDate = [datetime]::ParseExact($np.days[-1].date, 'yyyy-MM-dd', $null)
        if ($script:today -gt $lastDate) {
            $clamped = $script:today.AddDays(-(@($np.days).Count - 1)).ToString('yyyy-MM-dd')
            & $py (Join-Path $script:root '工具\build_plan.py') @commonArgs --start $clamped 2>&1 | Out-Null
        }
        if ($LASTEXITCODE -ne 0) {
            $script:statusText.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#C0392B')
            $script:statusText.Text = '重排失败'
            return
        }
        Write-Log ('重排计划: 每天 ' + $perDay)
        $script:win.Close()
        Start-Process -FilePath 'cmd.exe' -WindowStyle Hidden -ArgumentList '/c', ('timeout /t 2 /nobreak >nul & wscript.exe "' + (Join-Path $script:root 'launch.vbs') + '" preview')
    }

    function Start-NextDay {
        $planFile = Join-Path $script:root 'words.json'
        $plan = Get-Content $planFile -Raw -Encoding UTF8 | ConvertFrom-Json
        $curStart = [datetime]::ParseExact($plan.start, 'yyyy-MM-dd', $null)
        $newStart = $curStart.AddDays(-1)

        # 今天的打卡记录挪到昨天（它就是"下一天"的前一天）
        $todayFile = Join-Path $script:stateDir ($script:today.ToString('yyyy-MM-dd') + '.txt')
        $yestFile = Join-Path $script:stateDir ($script:today.AddDays(-1).ToString('yyyy-MM-dd') + '.txt')
        if (Test-Path $todayFile) {
            Copy-Item $todayFile $yestFile -Force
            Remove-Item $todayFile -Force
        }

        Fix-PlanDates -Dir $script:root -Start $newStart -Exam $script:planExam | Out-Null
        Write-Log ('提前一天: 计划起点改为 ' + $newStart.ToString('yyyy-MM-dd'))
        $script:closeTimer.Stop()
        $script:win.Close()
        Start-Process -FilePath 'cmd.exe' -WindowStyle Hidden -ArgumentList '/c', ('timeout /t 2 /nobreak >nul & wscript.exe "' + (Join-Path $script:root 'launch.vbs') + '" preview')
    }

    function Invoke-ResetDeck {
        $cfgFile = Join-Path $script:root 'config.json'
        $deckName = ''
        if (Test-Path $cfgFile) {
            try { $deckName = (Get-Content $cfgFile -Raw -Encoding UTF8 | ConvertFrom-Json).deck } catch { }
        }
        $deckPath = ''
        if ($deckName) { $deckPath = Join-Path $script:root ('词库\' + $deckName + '.csv') }
        if (-not (Test-Path $deckPath)) {
            $first = Get-ChildItem (Join-Path $script:root '词库') -Filter '*.csv' -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($first) { $deckPath = $first.FullName }
        }
        if (-not (Test-Path $deckPath)) {
            $script:statusText.Text = '找不到词库文件'
            $script:resetLayer.Visibility = 'Collapsed'
            return
        }
        Write-Log ('重置词库: ' + $deckPath)
        $script:resetLayer.Visibility = 'Collapsed'
        Invoke-DeckSwitch $deckPath
    }

    function Find-Python {
        $dsh = Join-Path $env:USERPROFILE '.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe'
        if (Test-Path $dsh) { return $dsh }
        $c = Get-Command python -ErrorAction SilentlyContinue
        if ($c) { return $c.Source }
        return $null
    }

    function Get-DeckFiles {
        $d = Join-Path $script:root '词库'
        if (-not (Test-Path $d)) { return @() }
        return @(Get-ChildItem $d -Filter '*.csv' -ErrorAction SilentlyContinue | Sort-Object Name)
    }

    function Get-ActiveDeck {
        $cfg = Join-Path $script:root 'config.json'
        if (Test-Path $cfg) {
            try { return (Get-Content $cfg -Raw -Encoding UTF8 | ConvertFrom-Json).deck } catch { }
        }
        return ''
    }

    function Update-DeckList {
        $list = $script:deckList
        $list.Children.Clear()
        $decks = @(Get-DeckFiles)
        if ($decks.Count -eq 0) {
            $list.Children.Add((New-Text $script:S.deckNone 12 '#999999' 'Normal' 4)) | Out-Null
            return
        }
        $active = Get-ActiveDeck
        foreach ($f in $decks) {
            $count = [Math]::Max(0, (Get-Content $f.FullName -Encoding UTF8).Count - 1)
            $label = $f.BaseName + '   （' + $count + ' 词）'
            if ($f.BaseName -eq $active) { $label = $label + '   · ' + $script:S.deckCurrent }
            $btn = New-Object Windows.Controls.Button
            $btn.Content = $label
            $btn.Style = $script:win.Resources['FlatBtn']
            $btn.HorizontalAlignment = 'Stretch'
            $btn.HorizontalContentAlignment = 'Left'
            $btn.Margin = New-Object Windows.Thickness(0, 3, 0, 3)
            if ($f.BaseName -eq $active) {
                $btn.Background = [Windows.Media.BrushConverter]::new().ConvertFromString('#1F4E79')
                $btn.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#FFFFFF')
            } else {
                $btn.Background = [Windows.Media.BrushConverter]::new().ConvertFromString('#EDF2F8')
                $btn.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#1F4E79')
            }
            $path = $f.FullName
            $btn.Add_Click({ param($s2, $e2) Show-DeckConfirm $path }.GetNewClosure())
            $list.Children.Add($btn) | Out-Null
        }
    }

    function Show-DeckConfirm([string]$path) {
        $list = $script:deckList
        $list.Children.Clear()
        $name = [System.IO.Path]::GetFileNameWithoutExtension($path)
        $list.Children.Add((New-Text ('切换到「' + $name + '」？') 13 '#1A1A1A' 'Bold' 2)) | Out-Null
        $row = New-Object Windows.Controls.StackPanel
        $row.Orientation = 'Horizontal'
        $row.Margin = New-Object Windows.Thickness(0, 14, 0, 0)

        $ok = New-Object Windows.Controls.Button
        $ok.Content = $script:S.deckConfirmYes
        $ok.Style = $script:win.Resources['FlatBtn']
        $ok.Background = [Windows.Media.BrushConverter]::new().ConvertFromString('#1F4E79')
        $ok.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#FFFFFF')
        $ok.Add_Click({ param($s2, $e2) Invoke-DeckSwitch $path }.GetNewClosure())

        $no = New-Object Windows.Controls.Button
        $no.Content = $script:S.deckConfirmNo
        $no.Style = $script:win.Resources['FlatBtn']
        $no.Background = [Windows.Media.BrushConverter]::new().ConvertFromString('#EDF2F8')
        $no.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#1F4E79')
        $no.Add_Click({ Update-DeckList })

        $row.Children.Add($ok) | Out-Null
        $row.Children.Add($no) | Out-Null
        $list.Children.Add($row) | Out-Null
    }

    function Invoke-DeckSwitch([string]$path) {
        try { Invoke-DeckSwitchInner $path }
        catch {
            Write-Log ('切换词库异常: ' + $_.Exception.Message)
            $script:statusText.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#C0392B')
            $script:statusText.Text = ('切换失败：' + $_.Exception.Message)
        }
    }

    function Invoke-DeckSwitchInner([string]$path) {
        $py = Find-Python
        if (-not $py) {
            $script:statusText.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#C0392B')
            $script:statusText.Text = '找不到 Python，无法切换词库'
            return
        }
        $gen = Join-Path $script:root '工具\build_plan.py'
        if (-not (Test-Path $gen)) {
            $script:statusText.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#C0392B')
            $script:statusText.Text = '安装版缺少生成器（工具\build_plan.py），无法重排计划'
            return
        }
        $script:statusText.Text = '正在生成…'
        $today = (Get-Date).ToString('yyyy-MM-dd')
        $exam = $script:planExam.ToString('yyyy-MM-dd')
        # 换词库 = 从新词库的第 1 个词重新开始；已打卡天数存在 state 里，不受影响
        $startStr = $script:today.ToString('yyyy-MM-dd')
        & $py (Join-Path $script:root '工具\build_plan.py') --deck $path --per-day $script:planPerDay `
            --intervals ($script:planIntervals -join ',') --start $startStr --exam $exam 2>&1 | Out-Null
        if ($LASTEXITCODE -ne 0) {
            $script:statusText.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#C0392B')
            $script:statusText.Text = '切换失败，看看词表有没有释义列'
            return
        }
        $cf = Join-Path $script:root 'config.json'
        $cc = $null
        if (Test-Path $cf) { try { $cc = Get-Content $cf -Raw -Encoding UTF8 | ConvertFrom-Json } catch { $cc = $null } }
        if (-not $cc) { $cc = New-Object psobject }
        $deckName = [System.IO.Path]::GetFileNameWithoutExtension($path)
        foreach ($pair in @(@('deck', $deckName), @('exam', $exam))) {
            if ($cc.PSObject.Properties[$pair[0]]) { $cc.($pair[0]) = $pair[1] }
            else { $cc | Add-Member -NotePropertyName $pair[0] -NotePropertyValue $pair[1] -Force }
        }
        ([string]($cc | ConvertTo-Json -Depth 4)) | Set-Content $cf -Encoding UTF8
        Write-Log ('切换词库: ' + $path)
        $script:win.Close()
        Start-Process -FilePath 'cmd.exe' -WindowStyle Hidden -ArgumentList '/c', ('timeout /t 2 /nobreak >nul & wscript.exe "' + (Join-Path $script:root 'launch.vbs') + '" preview')
    }

    # ---------- compact (narrow) layout ----------
    function Set-UiMode([string]$mode) {
        $m = 'normal'
        if ($mode -eq 'compact') { $m = 'compact' }
        $script:uiMode = $m
        $compact = ($m -eq 'compact')
        $win.FindName('MainScroll').Visibility = $(if ($compact) { 'Collapsed' } else { 'Visible' })
        $win.FindName('CompactScroll').Visibility = $(if ($compact) { 'Visible' } else { 'Collapsed' })
        $win.FindName('CompactFoot').Visibility = $(if ($compact) { 'Visible' } else { 'Collapsed' })
        $win.FindName('BtnPanel').Visibility = $(if ($compact) { 'Collapsed' } else { 'Visible' })
        $win.FindName('StatusText').Visibility = $(if ($compact) { 'Collapsed' } else { 'Visible' })
        $win.FindName('SubPill').Visibility = $(if ($compact) { 'Collapsed' } else { 'Visible' })
        $win.FindName('GearBtn').Visibility = $(if ($compact) { 'Collapsed' } else { 'Visible' })
        $win.FindName('DeckBtn').Visibility = $(if ($compact) { 'Collapsed' } else { 'Visible' })
        $win.FindName('PauseBtn').Visibility = $(if ($compact) { 'Collapsed' } else { 'Visible' })
        $win.FindName('NarrowBtn').Visibility = $(if ($compact) { 'Collapsed' } else { 'Visible' })
        $win.FindName('WideBtn').Visibility = $(if ($compact) { 'Visible' } else { 'Collapsed' })
        $title = $win.FindName('TitleText')
        $hdr = $win.FindName('HeaderGrid')
        if ($compact) {
            $win.Width = 400
            $title.FontSize = 13
            $title.Text = $script:titleShort
            $hdr.Margin = New-Object Windows.Thickness(16, 12, 14, 12)
        } else {
            $win.Width = 900
            $title.FontSize = 19
            $title.Text = $script:titleFull
            $hdr.Margin = New-Object Windows.Thickness(26, 18, 18, 18)
        }
        $win.UpdateLayout()
        if ($compact) {
            try {
                $wa = [Windows.SystemParameters]::WorkArea
                $win.Left = [math]::Max(0, ($wa.Width - $win.ActualWidth) / 2)
                $win.Top = 40
            } catch { }
        }
        Save-UiMode $m
        Write-Log ('ui mode: ' + $m)
    }

    function Save-UiMode([string]$mode) {
        try {
            $f = Join-Path $script:root 'config.json'
            $c = $null
            if (Test-Path $f) { $c = Get-Content $f -Raw -Encoding UTF8 | ConvertFrom-Json }
            if (-not $c) { $c = New-Object psobject }
            if ($c.PSObject.Properties['ui']) { $c.ui = $mode } else { $c | Add-Member -NotePropertyName ui -NotePropertyValue $mode -Force }
            ([string]($c | ConvertTo-Json -Depth 4)) | Set-Content $f -Encoding UTF8
        } catch { Write-Log ('save ui mode failed: ' + $_.Exception.Message) }
    }

    if (-not $isFinished) {
        $compactPanel = $win.FindName('CompactPanel')
        $compactPanel.Children.Add((New-Text ($S.secNew -f $day.words.Count) 10 '#1F4E79' 'Bold' 0)) | Out-Null
        foreach ($w in $day.words) {
            $row = New-Object Windows.Controls.Grid
            $row.Margin = New-Object Windows.Thickness(0, 5, 0, 5)
            $cA = New-Object Windows.Controls.ColumnDefinition
            $cA.Width = [Windows.GridLength]::new(0, [Windows.GridUnitType]::Auto)
            $cB = New-Object Windows.Controls.ColumnDefinition
            $cB.Width = [Windows.GridLength]::new(0, [Windows.GridUnitType]::Auto)
            $cC = New-Object Windows.Controls.ColumnDefinition
            $cC.Width = [Windows.GridLength]::new(1, [Windows.GridUnitType]::Star)
            $row.ColumnDefinitions.Add($cA) | Out-Null
            $row.ColumnDefinitions.Add($cB) | Out-Null
            $row.ColumnDefinitions.Add($cC) | Out-Null
            $head = New-Object Windows.Controls.StackPanel
            $head.Orientation = 'Horizontal'
            $head.Children.Add((New-Line $w.w 12.5 '#1A1A1A' 'Bold' 0)) | Out-Null
            if ($w.pos) { $head.Children.Add((New-Line $w.pos 9 '#999999' 'Normal' 4)) | Out-Null }
            [Windows.Controls.Grid]::SetColumn($head, 0)
            $row.Children.Add($head) | Out-Null
            if ($w.ipa) {
                $ipa = New-Line $w.ipa 10 '#5B7A99' 'Normal' 7
                $ipa.FontFamily = New-Object Windows.Media.FontFamily('Lucida Sans Unicode')
                [Windows.Controls.Grid]::SetColumn($ipa, 1)
                $row.Children.Add($ipa) | Out-Null
            }
            $cn = New-Line $w.cn 11 '#333333' 'Normal' 8
            $cn.ToolTip = $w.cn
            [Windows.Controls.Grid]::SetColumn($cn, 2)
            $row.Children.Add($cn) | Out-Null
            Add-CopyMenu $row (($w.w + '  ' + $w.cn).Trim())
            $compactPanel.Children.Add($row) | Out-Null
        }
    }

    function Invoke-DoneFlow([bool]$compact) {
        # spelling test first: only a full pass may check the day in.
        # never Hide() this window: hiding a ShowDialog window kills the modal loop.
        $script:win.Topmost = $false
        $script:win.Opacity = 0
        $script:savedLeft = $script:win.Left
        $script:savedTop = $script:win.Top
        $script:win.Left = -4000
        $script:win.Top = -4000
        $sp = Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\wscript.exe') -PassThru -Wait -ArgumentList @(
            ('"' + (Join-Path $script:root 'run-spell.vbs') + '"'), "$($script:dayIdx)")
        $script:win.Left = $script:savedLeft
        $script:win.Top = $script:savedTop
        $script:win.Opacity = 1
        $script:win.Topmost = $true
        if ($sp.ExitCode -eq 0) {
            Set-Content -Path (Join-Path $script:stateDir ($script:today.ToString('yyyy-MM-dd') + '.txt')) -Value 'done' -Encoding UTF8
            Write-Log $script:S.logDone
            $r = Sync-Workbook -Root $script:root -StateDir $script:stateDir
            Write-Log ('sync: ' + $r)
            if ($compact) {
                # 窄条模式：拼写通过后直接收起窗口，不再显示「已记录」横幅
                $script:win.Close()
                return
            }
            if ($r -eq 'ok') { Confirm-Saved $script:S.doneSaved } else { Confirm-Saved $script:S.doneSavedLater }
            return
        }
        $msg = $script:S.spellCancelled
        $fc = '#E6EEF7'
        if ($sp.ExitCode -eq 1) {
            Write-Log 'spelling round void'
            $msg = $script:S.spellVoid
            $fc = '#FFC9C9'
        } else {
            Write-Log 'spelling cancelled'
        }
        if ($compact) {
            $cs = $script:win.FindName('CompactStatus')
            $cs.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#C0392B')
            $cs.Text = $msg
        } else {
            $sub = $script:win.FindName('SubText')
            $sub.Text = $msg
            $sub.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString($fc)
        }
    }

    # ---------- buttons ----------
    if (-not $isFinished) {
        Add-Btn $S.btnDone '#1F4E79' '#FFFFFF' { Invoke-DoneFlow $false }
        Add-Btn $S.btnSnooze '#EDF2F8' '#1F4E79' {
            Write-Log $script:S.logSnooze
            Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden -ArgumentList @(
                '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command',
                ("Start-Sleep -Seconds 600; & '{0}'" -f $script:selfPath)
            )
            $script:win.Close()
        }
        Add-Btn $S.btnOpen '#EDF2F8' '#1F4E79' {
            $x = Get-ChildItem $script:root -Filter '*.xlsx' | Select-Object -First 1
            if ($x) { Start-Process $x.FullName }
        }
        Add-Btn $S.btnSkip '#FFFFFF' '#999999' {
            $script:confirmLayer.Visibility = 'Visible'
        }
        $cPanel = $win.FindName('CompactBtnPanel')
        Add-Btn $S.btnDone '#1F4E79' '#FFFFFF' { Invoke-DoneFlow $true } $cPanel | Out-Null
    }
    else {
        Add-Btn $S.btnOpen '#EDF2F8' '#1F4E79' {
            $x = Get-ChildItem $script:root -Filter '*.xlsx' | Select-Object -First 1
            if ($x) { Start-Process $x.FullName }
        }
        Add-Btn $S.btnOk '#1F4E79' '#FFFFFF' { $script:win.Close() }
        $win.FindName('CompactStatus').Text = $S.titleDone
        Add-Btn $S.btnOk '#1F4E79' '#FFFFFF' { $script:win.Close() } $win.FindName('CompactBtnPanel') | Out-Null
    }

    $confirmLayer.Add_MouseLeftButtonDown({
            param($s2, $e2)
            if ($e2.OriginalSource -eq $script:confirmLayer) { $script:confirmLayer.Visibility = 'Collapsed' }
        })
    $win.FindName('SetReset').Add_Click({ $script:resetLayer.Visibility = 'Visible' })
    $win.FindName('ResetYes').Add_Click({ Invoke-ResetDeck })
    $win.FindName('ResetNo').Add_Click({ $script:resetLayer.Visibility = 'Collapsed' })
    $resetLayer.Add_MouseLeftButtonDown({
            param($s2, $e2)
            if ($e2.OriginalSource -eq $script:resetLayer) { $script:resetLayer.Visibility = 'Collapsed' }
        })
    $win.FindName('SetSave').Add_Click({ Save-Target })
    $win.FindName('SetCancel').Add_Click({ $script:setLayer.Visibility = 'Collapsed' })
    $setLayer.Add_MouseLeftButtonDown({
            param($s2, $e2)
            if ($e2.OriginalSource -eq $script:setLayer) { $script:setLayer.Visibility = 'Collapsed' }
        })

    $deckBtn = $win.FindName('DeckBtn')
    $deckBtn.Add_MouseLeftButtonDown({
            param($s2, $e2)
            $e2.Handled = $true
            Update-DeckList
            $script:win.FindName('DeckFormat').Visibility = 'Collapsed'
            $script:deckLayer.Visibility = 'Visible'
        })
    $win.FindName('DeckHelp').Add_Click({
            $tb = $script:win.FindName('DeckFormat')
            if ($tb.Visibility -eq 'Collapsed') { $tb.Visibility = 'Visible' } else { $tb.Visibility = 'Collapsed' }
        })
    $win.FindName('DeckClose').Add_Click({ $script:deckLayer.Visibility = 'Collapsed' })
    $deckLayer.Add_MouseLeftButtonDown({
            param($s2, $e2)
            if ($e2.OriginalSource -eq $script:deckLayer) { $script:deckLayer.Visibility = 'Collapsed' }
        })
    $win.FindName('DeckAdd').Add_Click({
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
            $dlg = New-Object System.Windows.Forms.OpenFileDialog
            $dlg.Filter = '词表 (*.csv;*.txt)|*.csv;*.txt'
            if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $mgr = Join-Path $script:root 'deck-manager.ps1'
                # 不等待：导入跑在独立进程里，界面保持可用
                $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $mgr + '"'),
                    '-ImportFile', ('"' + $dlg.FileName + '"'))
                Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden -ArgumentList $argList | Out-Null
                $script:statusText.Text = '正在导入…'
                $script:deckTicks = 0
                if ($null -eq $script:deckPoll) {
                    $script:deckPoll = New-Object Windows.Threading.DispatcherTimer
                    $script:deckPoll.Interval = [TimeSpan]::FromMilliseconds(700)
                    $script:deckPoll.Add_Tick({
                            $script:deckTicks++
                            Update-DeckList
                            if ($script:deckTicks -ge 8) {
                                $script:deckPoll.Stop()
                                $script:statusText.Text = ''
                            }
                        })
                }
                $script:deckPoll.Start()
            }
        })

    $pauseBtn = $win.FindName('PauseBtn')
    $pauseBtn.Add_MouseLeftButtonDown({
            param($s2, $e2)
            $e2.Handled = $true
            $script:pauseLayer.Visibility = 'Visible'
        })
    $pauseLayer.Add_MouseLeftButtonDown({
            param($s2, $e2)
            if ($e2.OriginalSource -eq $script:pauseLayer) { $script:pauseLayer.Visibility = 'Collapsed' }
        })
    $win.FindName('PauseCancel').Add_Click({ $script:pauseLayer.Visibility = 'Collapsed' })
    $win.FindName('PauseD1').Add_Click({ Set-Pause 1 })
    $win.FindName('PauseD3').Add_Click({ Set-Pause 3 })
    $win.FindName('PauseD7').Add_Click({ Set-Pause 7 })
    $win.FindName('PauseAlways').Add_Click({ Set-Pause 0 })

    $win.FindName('BtnSkipNo').Add_Click({ $script:confirmLayer.Visibility = 'Collapsed' })
    $win.FindName('BtnSkipYes').Add_Click({
            Set-Content -Path (Join-Path $script:stateDir ($script:today.ToString('yyyy-MM-dd') + '.txt')) -Value 'skip' -Encoding UTF8
            Write-Log $script:S.logSkip
            $r = Sync-Workbook -Root $script:root -StateDir $script:stateDir
            Write-Log ('sync: ' + $r)
            $script:confirmLayer.Visibility = 'Collapsed'
            Show-UndoBanner $script:S.undoSkip
            $script:statusText.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#C0392B')
            $script:statusText.Text = $script:S.skippedNow
        })

    # ---------- restore saved ui mode ----------
    if ($script:uiModeSaved -eq 'compact') { Set-UiMode 'compact' }

    # ---------- show ----------
    Write-Log (($S.logShow -f $dayIdx))
    if ($Test) {
        if ($env:CET6_TEST_CONFIRM) { $win.FindName('ConfirmLayer').Visibility = 'Visible' }
        if ($env:CET6_TEST_PAUSE) { $win.FindName('PauseLayer').Visibility = 'Visible' }
        if ($env:CET6_TEST_DECK) { $win.FindName('DeckLayer').Visibility = 'Visible'; Update-DeckList }
        if ($env:CET6_TEST_DECKHELP) { $win.FindName('DeckLayer').Visibility = 'Visible'; Update-DeckList; $win.FindName('DeckFormat').Visibility = 'Visible' }
        if ($env:CET6_TEST_COMPACT) { Set-UiMode 'compact' }
        if ($env:CET6_TEST_SET) { Open-SetLayer }
        if ($env:CET6_TEST_RESET) { $win.FindName('ResetLayer').Visibility = 'Visible' }
        if ($env:CET6_TEST_REV) { foreach ($t in $flatList) { $t.Visibility = 'Collapsed' }; foreach ($t in $detailList) { $t.Visibility = 'Visible' }; $revToggle.Content = $S.revHide }
        if ($env:CET6_TEST_SCROLL) { $win.FindName('BodyPanel').Parent.ScrollToEnd() }
if ($env:CET6_TEST_UNDO) { Show-UndoBanner $S.undoSkip }
if ($env:CET6_TEST_PAUSEBANNER) { Show-PausedBanner ([string]::Format($S.pausedBanner, '11月9日')) }
        $timer = New-Object Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromMilliseconds(1000)
        $timer.Add_Tick({
            try {
                $rtb = New-Object Windows.Media.Imaging.RenderTargetBitmap -ArgumentList @(
                    [int]($script:win.ActualWidth * 2), [int]($script:win.ActualHeight * 2), 192, 192,
                    [Windows.Media.PixelFormats]::Pbgra32)
                $rtb.Render($script:win)
                $enc = New-Object Windows.Media.Imaging.PngBitmapEncoder
                $enc.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($rtb))
                $fs = [System.IO.File]::Create((Join-Path $script:root 'preview.png'))
                $enc.Save($fs)
                $fs.Close()
                Write-Log $script:S.logShot
            } catch {
                Write-Log ($script:S.logShotFail + $_.Exception.Message)
            }
            $timer.Stop()
            $script:win.Close()
        })
        $timer.Start()
        $win.ShowDialog() | Out-Null
        exit 0
    }

    $win.ShowDialog() | Out-Null
    Write-Log $S.logClosed
    exit 0
}
catch {
    Write-Log ($S.logError + $_.Exception.Message + " @ " + $_.InvocationInfo.PositionMessage)
    exit 1
}
