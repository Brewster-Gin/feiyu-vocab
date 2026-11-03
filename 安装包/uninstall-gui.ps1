# 肥鱼背单词 · 卸载界面（图形）
#   由 卸载.cmd / 开始菜单 / 「已安装的应用」调用；也可直接双击本文件所在目录的 uninstall-launch.vbs
param(
    [switch]$Test,
    [switch]$NoElevate,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

$dest = Split-Path -Parent $PSCommandPath
$uninstaller = Join-Path $dest 'uninstall-package.ps1'
$logOut = Join-Path $env:TEMP 'feiyu-uninstall-out.log'
$logErr = Join-Path $env:TEMP 'feiyu-uninstall-err.log'

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin -and -not $NoElevate -and -not $Test -and -not $DryRun) {
    try {
        $argLine = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $PSCommandPath + '" -NoElevate'
        Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList $argLine -WindowStyle Hidden -Wait | Out-Null
        exit 0
    } catch {
        # 用户拒绝 UAC：留在非管理员模式继续
    }
}

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="肥鱼背单词 · 卸载" Width="700" Height="604"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
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
        <TextBlock Text="肥鱼背单词 · 卸载" Foreground="#FFFFFF" FontSize="19" FontWeight="Bold"/>
        <TextBlock Name="HeaderSub" Text="版本 1.4" Foreground="#CFE2F3" FontSize="11.5" Margin="0,6,0,0"/>
      </StackPanel>
    </Border>

    <StackPanel Grid.Row="1" Margin="28,18,28,0">
      <TextBlock Text="程序目录" FontSize="12.5" FontWeight="Bold" Foreground="#1A1A1A"/>
      <TextBlock Name="DestText" FontSize="11" Foreground="#666666" Margin="0,6,0,0" TextWrapping="Wrap"/>

      <TextBlock Text="选择卸载方式" FontSize="12.5" FontWeight="Bold" Foreground="#1A1A1A" Margin="0,16,0,8"/>

      <Border Name="CardKeep" CornerRadius="10" Padding="14,12" Cursor="Hand"
              Background="#FFFFFF" BorderBrush="#CBD5E0" BorderThickness="1">
        <StackPanel>
          <TextBlock Name="CardKeepTitle" Text="只卸载程序，保留数据" FontSize="13" FontWeight="Bold" Foreground="#1A1A1A"/>
          <TextBlock Text="保留打卡表、词库、学习记录和计划设置；以后重新安装可以直接接着用。"
                     FontSize="11" Foreground="#7A7A7A" Margin="0,5,0,0" TextWrapping="Wrap"/>
        </StackPanel>
      </Border>

      <Border Name="CardAll" CornerRadius="10" Padding="14,12" Cursor="Hand" Margin="0,10,0,0"
              Background="#FFFFFF" BorderBrush="#CBD5E0" BorderThickness="1">
        <StackPanel>
          <TextBlock Name="CardAllTitle" Text="全部删除" FontSize="13" FontWeight="Bold" Foreground="#1A1A1A"/>
          <TextBlock Text="连打卡表、词库、学习记录一起删掉，不可恢复。"
                     FontSize="11" Foreground="#7A7A7A" Margin="0,5,0,0" TextWrapping="Wrap"/>
        </StackPanel>
      </Border>

      <TextBlock Text="卸载日志" FontSize="12.5" FontWeight="Bold" Foreground="#1A1A1A" Margin="0,16,0,6"/>
      <Border BorderBrush="#DDE3EA" BorderThickness="1" CornerRadius="6" Background="#FFFFFF" Height="196">
        <TextBox Name="LogBox" IsReadOnly="True" BorderThickness="0" Padding="10,8"
                 FontFamily="Consolas" FontSize="10.5" Foreground="#333333"
                 VerticalScrollBarVisibility="Auto" TextWrapping="Wrap" Background="Transparent"/>
      </Border>
    </StackPanel>

    <Border Grid.Row="2" Background="#FFFFFF" BorderBrush="#E6EAF0" BorderThickness="0,1,0,0" Margin="0,14,0,0">
      <Grid Margin="28,14,28,16">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock Name="StatusText" Grid.Column="0" FontSize="11.5" Foreground="#666666"
                   VerticalAlignment="Center" TextWrapping="Wrap" Text="请选择一种方式"/>
        <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
          <Button Name="BtnClose" Content="完成" Style="{StaticResource PrimaryBtn}" Visibility="Collapsed" Margin="0,0,10,0"/>
          <Button Name="BtnCancel" Content="取消" Style="{StaticResource GhostBtn}" Margin="0,0,10,0"/>
          <Button Name="BtnGo" Content="开始卸载" Style="{StaticResource PrimaryBtn}"/>
        </StackPanel>
      </Grid>
    </Border>
  </Grid>
</Window>
'@

$win = [Windows.Markup.XamlReader]::Parse($xaml)
$script:mode = 'keep'
$script:running = $false

$statusText = $win.FindName('StatusText')
$logBox = $win.FindName('LogBox')
$btnGo = $win.FindName('BtnGo')
$btnCancel = $win.FindName('BtnCancel')
$btnClose = $win.FindName('BtnClose')
$cardKeep = $win.FindName('CardKeep')
$cardAll = $win.FindName('CardAll')

$win.FindName('DestText').Text = $dest
$win.FindName('HeaderSub').Text = '版本 1.4    卸载不会影响你的打卡表（除非选「全部删除」）'

function Brush([string]$hex) { [Windows.Media.BrushConverter]::new().ConvertFromString($hex) }

function Set-Cardlook {
    if ($script:mode -eq 'keep') {
        $cardKeep.Background = Brush '#EAF2FB'; $cardKeep.BorderBrush = Brush '#2E75B6'
        $win.FindName('CardKeepTitle').Foreground = Brush '#1F4E79'
        $cardAll.Background = Brush '#FFFFFF'; $cardAll.BorderBrush = Brush '#CBD5E0'
        $win.FindName('CardAllTitle').Foreground = Brush '#1A1A1A'
        $btnGo.Content = '卸载程序'
        $statusText.Text = '将删除程序文件，保留打卡表和学习记录'
    } else {
        $cardAll.Background = Brush '#FDECEA'; $cardAll.BorderBrush = Brush '#C0392B'
        $win.FindName('CardAllTitle').Foreground = Brush '#A93226'
        $cardKeep.Background = Brush '#FFFFFF'; $cardKeep.BorderBrush = Brush '#CBD5E0'
        $win.FindName('CardKeepTitle').Foreground = Brush '#1A1A1A'
        $btnGo.Content = '全部删除'
        $statusText.Text = '将连打卡表和学习记录一起删除，不可恢复'
    }
}

function Select-Card([string]$m) {
    if ($script:running) { return }
    $script:mode = $m
    Set-Cardlook
}

$cardKeep.Add_MouseLeftButtonDown({ Select-Card 'keep' })
$cardAll.Add_MouseLeftButtonDown({ Select-Card 'all' })
Set-Cardlook

function Read-Shared([string]$path) {
    if (-not (Test-Path $path)) { return '' }
    try {
        $fs = [IO.File]::Open($path, 'Open', 'Read', 'ReadWrite')
        $buf = New-Object byte[] $fs.Length
        [void]$fs.Read($buf, 0, $buf.Length)
        $fs.Close()
        try { $txt = (New-Object Text.UTF8Encoding($false, $true)).GetString($buf) }
        catch { $txt = [Text.Encoding]::GetEncoding(936).GetString($buf) }
        return $txt
    } catch { return '' }
}

$script:shown = 0
function Update-Log {
    $txt = Read-Shared $logOut
    if ($txt.Length -gt $script:shown) {
        $add = $txt.Substring($script:shown)
        $script:shown = $txt.Length
        $logBox.AppendText($add.Replace("`n", "`r`n"))
        $logBox.ScrollToEnd()
    }
}

$btnCancel.Add_Click({
        if ($script:running) { return }
        $win.Close()
    })

$btnClose.Add_Click({ $win.Close() })

$script:timer = New-Object Windows.Threading.DispatcherTimer
$script:timer.Interval = [TimeSpan]::FromMilliseconds(260)

$btnGo.Add_Click({
        if ($script:running) { return }

        if ($script:mode -eq 'all') {
            $r = [Windows.MessageBox]::Show(
                "将删除整个程序文件夹，打卡表和学习记录都不保留。`n`n确定继续吗？",
                '确认全部删除', 'YesNo', 'Warning')
            if ($r -ne 'Yes') { return }
        }

        $script:running = $true
        $btnGo.IsEnabled = $false
        $cardKeep.IsHitTestVisible = $false
        $cardAll.IsHitTestVisible = $false
        $logBox.Clear()
        $script:shown = 0
        $statusText.Text = '正在卸载…'

        foreach ($f in @($logOut, $logErr)) { try { Remove-Item $f -Force -ErrorAction SilentlyContinue } catch { } }

        if ($DryRun) {
            $fake = @(
                '==== 肥鱼背单词 · 卸载 ====',
                ('程序目录: ' + $dest),
                '[OK] 已结束 1 个占用进程',
                '[OK] 已删除计划任务',
                '[OK] 已删除桌面图标: 肥鱼背单词.lnk',
                '[OK] 已删除开始菜单项',
                '[OK] 已从「已安装的应用」列表移除',
                '[OK] 已删除 14 项程序文件',
                '[OK] 已保留：打卡表 xlsx、词库、学习记录、计划设置',
                '',
                '卸载完成。'
            )
            Set-Content -Path $logOut -Value ($fake -join "`n") -Encoding UTF8
        } else {
            $swArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $uninstaller + '"'),
                '-Silent', '-NoElevate')
            if ($script:mode -eq 'keep') { $swArgs += '-KeepData' }
            try {
                $script:proc = Start-Process -FilePath 'powershell.exe' -ArgumentList $swArgs `
                    -WindowStyle Hidden -PassThru `
                    -RedirectStandardOutput $logOut -RedirectStandardError $logErr
            } catch {
                $script:running = $false
                $btnGo.IsEnabled = $true
                $cardKeep.IsHitTestVisible = $true
                $cardAll.IsHitTestVisible = $true
                $statusText.Text = '启动卸载失败：' + $_.Exception.Message
                return
            }
        }

        $script:ticks = 0
        $script:timer.Start()
    })

$script:timer.Add_Tick({
        $script:ticks++
        Update-Log
        $alive = $false
        if ($script:proc) { $alive = -not $script:proc.HasExited }
        if ($DryRun) { $alive = $script:ticks -lt 12 }
        if ($alive) { return }

        $script:timer.Stop()
        Update-Log
        $script:running = $false
        $code = 0
        if ($script:proc) { try { $code = $script:proc.ExitCode } catch { $code = -1 } }
        $uniLog = ''
        try { $uniLog = [IO.File]::ReadAllText((Join-Path $env:TEMP 'feiyu-uninstall-out.log'), [Text.Encoding]::UTF8) } catch { }
        $uniDone = ($DryRun -or $uniLog -match '卸载完成')

        if ($uniDone) {
            $statusText.Text = if ($script:mode -eq 'keep') { '卸载完成，打卡表和学习记录保留在原文件夹' } else { '卸载完成' }
            $btnGo.Visibility = 'Collapsed'
            $btnCancel.Visibility = 'Collapsed'
            $btnClose.Visibility = 'Visible'
        } else {
            $statusText.Text = ('卸载返回错误码 ' + $code + '，详见 %TEMP%\feiyu-uninstall-err.log')
            $btnGo.IsEnabled = $true
            $btnGo.Content = '重试'
            $cardKeep.IsHitTestVisible = $true
            $cardAll.IsHitTestVisible = $true
        }
    })

$win.Add_ContentRendered({
        if ($DryRun) { $btnGo.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent))) }
    })

if ($Test) {
    $shot = if ($env:CET6_TEST_OUT) { $env:CET6_TEST_OUT } else { Join-Path $dest 'uninstall-preview.png' }
    $win.Show()
    $win.UpdateLayout()
    Start-Sleep -Milliseconds 400
    $bmp = New-Object Windows.Media.Imaging.RenderTargetBitmap(
        [int]$win.ActualWidth, [int]$win.ActualHeight, 96, 96, [Windows.Media.PixelFormats]::Pbgra32)
    $bmp.Render($win)
    $enc = New-Object Windows.Media.Imaging.PngBitmapEncoder
    $enc.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bmp))
    $fsImg = [IO.File]::Create($shot)
    $enc.Save($fsImg)
    $fsImg.Close()
    $win.Close()
} else {
    [void]$win.ShowDialog()
}
