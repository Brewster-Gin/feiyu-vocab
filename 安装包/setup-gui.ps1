param(
    [switch]$Test,
    [switch]$DryRun,
    [switch]$NoElevate,
    [string]$Path = ''
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$installer = Join-Path $root 'install-package.ps1'
$logFile = Join-Path $env:TEMP 'feiyu-install.log'

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

$regPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\FeiyuVocab'
$prev = ''
try { $prev = (Get-ItemProperty -Path $regPath -Name InstallLocation -ErrorAction Stop).InstallLocation } catch { }
$defaultDest = if ($prev) { $prev } else { Join-Path $env:LOCALAPPDATA 'CET6Plan' }
$dest0 = if ($Path) { $Path } else { $defaultDest }

# ---------- 提权（一次，启动时） ----------
if (-not $isAdmin -and -not $NoElevate -and -not $Test -and -not $DryRun) {
    try {
        $argLine = '-NoProfile -ExecutionPolicy Bypass -File "' + $PSCommandPath + '" -NoElevate -Path "' + $dest0 + '"'
        Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList $argLine -Wait | Out-Null
        exit 0
    } catch {
        # 用户拒绝 UAC：留在非管理员模式继续
    }
}

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="肥鱼背单词 · 安装" Width="700" Height="648"
        WindowStartupLocation="CenterScreen" ResizeMode="CanMinimize"
        Background="#F5F7FA" FontFamily="Microsoft YaHei UI">
  <Window.Resources>
    <Style x:Key="PrimaryBtn" TargetType="Button">
      <Setter Property="Foreground" Value="#FFFFFF"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="#2E75B6" CornerRadius="8" Padding="22,10">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#3D8BD1"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="bd" Property="Background" Value="#B9CBDB"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="GhostBtn" TargetType="Button">
      <Setter Property="Foreground" Value="#1F4E79"/>
      <Setter Property="FontSize" Value="12.5"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="#EDF2F7" CornerRadius="8" Padding="18,9">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#E1EAF4"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="bd" Property="Opacity" Value="0.5"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0">
      <Border.Background>
        <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
          <GradientStop Color="#1F4E79" Offset="0"/>
          <GradientStop Color="#2E75B6" Offset="1"/>
        </LinearGradientBrush>
      </Border.Background>
      <StackPanel Margin="28,20,28,18">
        <TextBlock Text="肥鱼背单词 · 安装" Foreground="#FFFFFF" FontSize="19" FontWeight="Bold"/>
        <TextBlock Text="版本 1.3    装好后每天打开电脑自动提醒" Foreground="#CFE2F3" FontSize="11.5" Margin="0,6,0,0"/>
      </StackPanel>
    </Border>

    <StackPanel Grid.Row="1" Margin="28,20,28,0">
      <TextBlock Text="安装位置" FontSize="12.5" FontWeight="Bold" Foreground="#1A1A1A"/>
      <Grid Margin="0,8,0,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBox Name="DestBox" Grid.Column="0" FontSize="12.5" Padding="9,7"
                 VerticalContentAlignment="Center" BorderBrush="#CBD5E0"/>
        <Button Name="BrowseBtn" Grid.Column="1" Content="浏览…" Margin="10,0,0,0"
                Width="96" Style="{StaticResource GhostBtn}"/>
      </Grid>
      <TextBlock Name="DestHint" FontSize="10.5" Foreground="#8A8A8A" Margin="0,6,0,0" TextWrapping="Wrap"/>

      <StackPanel Orientation="Horizontal" Margin="0,18,0,0">
        <CheckBox Name="CbDesktop" Content="桌面图标" IsChecked="True" FontSize="12" Foreground="#333333"/>
        <CheckBox Name="CbStart" Content="开始菜单" IsChecked="True" FontSize="12" Foreground="#333333" Margin="24,0,0,0"/>
        <CheckBox Name="CbAuto" Content="开机 / 解锁 / 20:30 自动提醒" IsChecked="True" FontSize="12" Foreground="#333333" Margin="24,0,0,0"/>
      </StackPanel>
      <TextBlock Name="AdminHint" FontSize="10.5" Foreground="#B26A00" Margin="0,8,0,0" TextWrapping="Wrap"/>

      <TextBlock FontSize="10.5" Foreground="#8A8A8A" Margin="0,14,0,0" TextWrapping="Wrap"
                 Text="装好后打开弹窗，点日期右边的小齿轮可以设置考试日期、每天背几个词、切换词库。"/>

      <TextBlock Text="安装日志" FontSize="12.5" FontWeight="Bold" Foreground="#1A1A1A" Margin="0,16,0,6"/>
      <Border BorderBrush="#DDE3EA" BorderThickness="1" CornerRadius="6" Background="#FFFFFF" Height="238">
        <TextBox Name="LogBox" IsReadOnly="True" BorderThickness="0" Padding="10,8"
                 FontFamily="Consolas" FontSize="10.5" Foreground="#333333"
                 VerticalScrollBarVisibility="Auto" TextWrapping="Wrap" Background="Transparent"/>
      </Border>
    </StackPanel>

    <Border Grid.Row="2" Background="#FFFFFF" BorderBrush="#E6EAF0" BorderThickness="0,1,0,0" Margin="0,16,0,0">
      <Grid Margin="28,14,28,16">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock Name="StatusText" Grid.Column="0" FontSize="11.5" Foreground="#666666"
                   VerticalAlignment="Center" TextWrapping="Wrap" Text="准备就绪"/>
        <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
          <Button Name="BtnOpen" Content="打开程序" Style="{StaticResource GhostBtn}" Visibility="Collapsed" Margin="0,0,10,0"/>
          <Button Name="BtnClose" Content="完成" Style="{StaticResource PrimaryBtn}" Visibility="Collapsed" Margin="0,0,10,0"/>
          <Button Name="BtnCancel" Content="取消" Style="{StaticResource GhostBtn}" Margin="0,0,10,0"/>
          <Button Name="BtnInstall" Content="开始安装" Style="{StaticResource PrimaryBtn}"/>
        </StackPanel>
      </Grid>
    </Border>
  </Grid>
</Window>
'@

$win = [Windows.Markup.XamlReader]::Parse($xaml)

function Get-Named([string]$name) { $win.FindName($name) }

$destBox = Get-Named 'DestBox'
$browseBtn = Get-Named 'BrowseBtn'
$destHint = Get-Named 'DestHint'
$cbDesktop = Get-Named 'CbDesktop'
$cbStart = Get-Named 'CbStart'
$cbAuto = Get-Named 'CbAuto'
$adminHint = Get-Named 'AdminHint'
$logBox = Get-Named 'LogBox'
$statusText = Get-Named 'StatusText'
$btnInstall = Get-Named 'BtnInstall'
$btnCancel = Get-Named 'BtnCancel'
$btnOpen = Get-Named 'BtnOpen'
$btnClose = Get-Named 'BtnClose'

$destBox.Text = $dest0
$destHint.Text = '默认位置：' + $defaultDest
if ($prev) { $destHint.Text = '检测到已安装在：' + $prev + '（继续会覆盖升级）' }

if (-not $isAdmin) {
    $adminHint.Text = '当前没有管理员权限：装完只能做「开机提醒」，「解锁」和「每天 20:30」需要管理员。想全开就重开安装包，在 UAC 里点「是」。'
}

$script:running = $false
$script:proc = $null
$script:timer = $null
$script:shown = 0
$script:outFile = Join-Path $env:TEMP 'feiyu-install-out.log'
$script:errFile = Join-Path $env:TEMP 'feiyu-install-err.log'
$script:fakeIdx = 0
$script:fakeLines = @()

function Read-Shared([string]$path) {
    if (-not (Test-Path $path)) { return '' }
    try {
        $fs = [IO.File]::Open($path, 'Open', 'Read', 'ReadWrite')
        $buf = New-Object byte[] $fs.Length
        [void]$fs.Read($buf, 0, $buf.Length)
        $fs.Close()
        try {
            $txt = (New-Object Text.UTF8Encoding($false, $true)).GetString($buf)
        } catch {
            $txt = [Text.Encoding]::GetEncoding(936).GetString($buf)
        }
        return $txt
    } catch { return '' }
}

function Append-Log([string]$text) {
    if (-not $text) { return }
    $logBox.AppendText($text + "`r`n")
    $logBox.ScrollToEnd()
}

function Set-Running([bool]$on) {
    $script:running = $on
    $btnInstall.IsEnabled = (-not $on)
    $destBox.IsEnabled = (-not $on)
    $browseBtn.IsEnabled = (-not $on)
    $cbDesktop.IsEnabled = (-not $on)
    $cbStart.IsEnabled = (-not $on)
    $cbAuto.IsEnabled = (-not $on)
    $btnCancel.IsEnabled = (-not $on)
}

function Show-Done([bool]$ok, [string]$dest) {
    Set-Running $false
    $btnInstall.Visibility = 'Collapsed'
    $btnCancel.Visibility = 'Collapsed'
    $btnClose.Visibility = 'Visible'
    if ($ok) {
        $btnOpen.Visibility = 'Visible'
        $statusText.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#2E7D4F')
        $statusText.Text = '安装完成 · ' + $dest
        Append-Log ''
        Append-Log '==== 安装完成 ===='
        Append-Log ('安装位置: ' + $dest)
    } else {
        $statusText.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#C0392B')
        $statusText.Text = '安装失败，日志在 ' + $logFile
    }
}

function Complete-Install {
    $code = 0
    try { $code = $script:proc.ExitCode } catch { $code = -1 }
    $err = Read-Shared $script:errFile
    if ($err) { Append-Log ('[stderr] ' + $err.Trim()) }
    $dest = $destBox.Text.Trim()
    if ($code -eq 0) {
        try { Copy-Item $script:outFile $logFile -Force } catch { }
        Show-Done $true $dest
    } else {
        Show-Done $false $dest
        Append-Log ('[!!] 安装脚本退出码 ' + $code)
    }
}

function Start-Poll {
    $script:timer = New-Object Windows.Threading.DispatcherTimer
    $script:timer.Interval = [TimeSpan]::FromMilliseconds(260)
    $script:timer.Add_Tick({
            if ($script:proc -and $script:proc.HasExited) {
                $script:timer.Stop()
                Complete-Install
                return
            }
            $txt = Read-Shared $script:outFile
            if ($txt) {
                $lines = $txt -split "`r?`n"
                if ($lines.Count -gt $script:shown) {
                    for ($i = $script:shown; $i -lt $lines.Count; $i++) {
                        if ($lines[$i].Trim()) { Append-Log $lines[$i] }
                    }
                    $script:shown = $lines.Count
                }
            }
        })
    $script:timer.Start()
}

function Start-Install {
    if ($script:running) { return }
    $dest = $destBox.Text.Trim().Trim('"')
    if (-not $dest) { $dest = $defaultDest }
    if (-not [IO.Path]::IsPathRooted($dest)) { $dest = Join-Path (Get-Location) $dest }
    $destBox.Text = $dest

    if (Test-Path $dest -PathType Leaf) {
        $statusText.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#C0392B')
        $statusText.Text = '这个路径是一个文件，请换成文件夹'
        return
    }
    if (Test-Path $dest) {
        $items = @(Get-ChildItem $dest -Force -ErrorAction SilentlyContinue)
        if ($items.Count -gt 0) {
            $msg = '目标文件夹里已经有 ' + $items.Count + ' 个文件：' + "`r`n" + $dest + "`r`n`r`n继续安装会覆盖同名文件，确定吗？"
            $ans = [Windows.MessageBox]::Show($msg, '肥鱼背单词', 'YesNo', 'Question')
            if ($ans -ne 'Yes') { return }
        }
    }

    $logBox.Clear()
    $script:shown = 0
    if (Test-Path $script:outFile) { Remove-Item $script:outFile -Force -ErrorAction SilentlyContinue }
    if (Test-Path $script:errFile) { Remove-Item $script:errFile -Force -ErrorAction SilentlyContinue }

    Set-Running $true
    $statusText.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#1F4E79')
    $statusText.Text = '正在安装…'

    if ($DryRun) {
        $script:fakeLines = @(
            '==== 肥鱼背单词 · 安装 ====',
            ('安装位置: ' + $dest),
            ('开始日期: ' + (Get-Date).ToString('yyyy-MM-dd')),
            ('管理员  : ' + $(if ($isAdmin) { '是（含解锁/定时提醒）' } else { '否（只有开机提醒）' })),
            '[OK] 生成器已就位（可切换词库/改每日词数）',
            '[OK] 词库已就位（1 套）',
            '[OK] 程序文件已复制',
            ('[OK] 计划排到 ' + (Get-Date).ToString('yyyy-MM-dd')),
            '[OK] 图标已就位',
            '[OK] 桌面图标 + 开始菜单（含「卸载肥鱼背单词」）已创建',
            '[OK] 已注册提醒任务（登录 / 解锁 / 每天 20:30）',
            '[OK] 已登记到「设置 → 已安装的应用」'
        )
        $script:fakeIdx = 0
        $fakeTimer = New-Object Windows.Threading.DispatcherTimer
        $fakeTimer.Interval = [TimeSpan]::FromMilliseconds(200)
        $fakeTimer.Add_Tick({
                if ($script:fakeIdx -lt $script:fakeLines.Count) {
                    Append-Log $script:fakeLines[$script:fakeIdx]
                    $script:fakeIdx++
                } else {
                    $fakeTimer.Stop()
                    Show-Done $true $dest
                }
            })
        $fakeTimer.Start()
        return
    }

    $swArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $installer + '"', '-NoElevate'),
        '-Path', ('"' + $dest + '"'), '-Silent')
    if (-not $cbDesktop.IsChecked) { $swArgs += '-NoDesktop' }
    if (-not $cbStart.IsChecked) { $swArgs += '-NoStartMenu' }
    if (-not $cbAuto.IsChecked) { $swArgs += '-NoAutoStart' }

    try {
        $script:proc = Start-Process -FilePath 'powershell.exe' -ArgumentList ($swArgs -join ' ') `
            -WindowStyle Hidden -PassThru -RedirectStandardOutput $script:outFile -RedirectStandardError $script:errFile
        Start-Poll
    } catch {
        Set-Running $false
        $statusText.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString('#C0392B')
        $statusText.Text = '启动安装失败：' + $_.Exception.Message
    }
}

$browseBtn.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = '选择安装位置'
        $dlg.ShowNewFolderButton = $true
        $cur = $destBox.Text.Trim()
        if ($cur -and (Test-Path $cur)) { $dlg.SelectedPath = $cur }
        elseif (Test-Path $defaultDest) { $dlg.SelectedPath = $defaultDest }
        else { $dlg.SelectedPath = $env:LOCALAPPDATA }
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $destBox.Text = $dlg.SelectedPath
        }
    })

$destBox.Add_TextChanged({
        $t = $destBox.Text.Trim().Trim('"')
        if ($t -and (Test-Path $t -PathType Container)) {
            $n = @(Get-ChildItem $t -Force -ErrorAction SilentlyContinue).Count
            if ($n -gt 0) { $destHint.Text = '该文件夹已有 ' + $n + ' 个文件，继续会覆盖同名文件' }
            else { $destHint.Text = '这个文件夹是空的' }
        } else {
            $destHint.Text = '默认位置：' + $defaultDest
        }
    })

$btnInstall.Add_Click({ Start-Install })
$btnCancel.Add_Click({ $win.Close() })
$btnClose.Add_Click({ $win.Close() })

$btnOpen.Add_Click({
        $dest = $destBox.Text.Trim()
        $vbs = Join-Path $dest 'launch.vbs'
        if (Test-Path $vbs) {
            Start-Process -FilePath "$env:SystemRoot\System32\wscript.exe" -ArgumentList ('"' + $vbs + '" preview')
        } else {
            Start-Process explorer.exe $dest
        }
    })

$win.Add_Closing({
        param($s, $e)
        if ($script:running -and $script:proc -and -not $script:proc.HasExited) {
            $ans = [Windows.MessageBox]::Show('安装还没结束，现在关闭会中断安装，确定吗？', '肥鱼背单词', 'YesNo', 'Warning')
            if ($ans -ne 'Yes') { $e.Cancel = $true; return }
            try { $script:proc.Kill() } catch { }
        }
    })

if ($Test) {
    $shotPath = if ($env:CET6_TEST_OUT) { $env:CET6_TEST_OUT } else { Join-Path $root 'installer-preview.png' }
    Append-Log '==== 肥鱼背单词 · 安装 ===='
    Append-Log ('安装位置: ' + $dest0)
    Append-Log '[OK] 生成器已就位（可切换词库/改每日词数）'
    Append-Log '[OK] 词库已就位（1 套）'
    Append-Log '[OK] 程序文件已复制'
    Append-Log '[OK] 已注册提醒任务（登录 / 解锁 / 每天 20:30）'
    $statusText.Text = '正在安装…'
    $t = New-Object Windows.Threading.DispatcherTimer
    $t.Interval = [TimeSpan]::FromMilliseconds(900)
    $t.Add_Tick({
            try {
                $rtb = New-Object Windows.Media.Imaging.RenderTargetBitmap -ArgumentList @(
                    [int]($win.ActualWidth * 2), [int]($win.ActualHeight * 2), 192, 192,
                    [Windows.Media.PixelFormats]::Pbgra32)
                $rtb.Render($win)
                $enc = New-Object Windows.Media.Imaging.PngBitmapEncoder
                $enc.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($rtb))
                $fs = [IO.File]::Create($shotPath)
                $enc.Save($fs); $fs.Close()
            } catch {
                [IO.File]::WriteAllText((Join-Path $env:TEMP 'installer-preview-err.txt'), $_.Exception.ToString())
            }
            $t.Stop()
            $win.Close()
        })
    $t.Start()
}

$win.ShowDialog() | Out-Null
