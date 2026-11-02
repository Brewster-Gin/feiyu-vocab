# 拼写测试：给中文，默写英文。全部拼对才允许打卡。
# 由 reminder.ps1 调用：powershell -File spell.ps1 -Day 3
# 退出码：0 = 全部通过；2 = 取消
param(
    [int]$Day = 1,
    [switch]$Test
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$logFile = Join-Path $root 'reminder.log'

function Write-Log([string]$msg) {
    try { Add-Content -Path $logFile -Value ("[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg) -Encoding UTF8 } catch { }
}

try {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
    [System.Windows.Media.RenderOptions]::ProcessRenderMode = [System.Windows.Interop.RenderMode]::SoftwareOnly

    $plan = Get-Content (Join-Path $root 'words.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($Day -lt 1 -or $Day -gt @($plan.days).Count) { exit 2 }
    $newWords = @($plan.days[$Day - 1].words)

    # 从当天的复习组里随机抽几个一起考
    $reviewPool = @()
    $intervals = @($plan.intervals)
    $nDays = @($plan.days).Count
    foreach ($iv in $intervals) {
        $src = $Day - [int]$iv - 1
        if ($src -ge 0 -and $src -lt $nDays) {
            foreach ($rw in @($plan.days[$src].words)) { $reviewPool += $rw }
        }
    }
    $reviewPick = @()
    if ($reviewPool.Count -gt 0) {
        $take = [Math]::Min(5, $reviewPool.Count)
        $reviewPick = @($reviewPool | Get-Random -Count $take)
    }

    $words = @()
    foreach ($it in $newWords) {
        $words += [pscustomobject]@{ w = $it.w; pos = $it.pos; cn = $it.cn; ipa = $it.ipa; rev = $false }
    }
    foreach ($it in $reviewPick) {
        $words += [pscustomobject]@{ w = $it.w; pos = $it.pos; cn = $it.cn; ipa = $it.ipa; rev = $true }
    }
    # 新词在前、复习词在后，但整体打散一次，避免固定顺序
    if ($words.Count -gt 1) { $words = @($words | Get-Random -Count $words.Count) }
    $total = $words.Count

    $script:idx = 0
    $script:wrong = 0
    $script:firstTry = 0
    $script:passed = $false
    $script:revealed = $false
    $script:busy = $false
    $script:round = 1
    $script:void = $false

    $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        ResizeMode="NoResize" SizeToContent="Height" Width="840"
        WindowStartupLocation="CenterScreen" ShowInTaskbar="False" Topmost="True"
        FontFamily="Microsoft YaHei UI">
  <Window.Resources>
    <Style x:Key="FlatBtn" TargetType="Button">
      <Setter Property="Padding" Value="16,9"/>
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
  <Border CornerRadius="16" Background="#FFFFFF" Margin="18">
    <Border.LayoutTransform>
      <ScaleTransform ScaleX="1.5" ScaleY="1.5"/>
    </Border.LayoutTransform>
    <Border.Effect>
      <DropShadowEffect BlurRadius="28" ShadowDepth="6" Opacity="0.28" Color="#000000"/>
    </Border.Effect>
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
        <StackPanel Margin="26,16,26,16">
          <TextBlock Name="TitleText" Text="拼写测试 · 全部拼对才能打卡" Foreground="#FFFFFF"
                     FontSize="17" FontWeight="Bold"/>
          <TextBlock Name="ProgressText" Text="" Foreground="#CFE2F3" FontSize="12" Margin="0,5,0,0"/>
        </StackPanel>
      </Border>
      <StackPanel Grid.Row="1" Margin="26,22,26,8">
        <TextBlock Name="PromptText" Text="" FontSize="26" FontWeight="Bold" Foreground="#1A1A1A"
                   TextWrapping="Wrap"/>
        <TextBlock Name="PosText" Text="" FontSize="12" Foreground="#999999" Margin="0,8,0,0"/>
        <TextBox Name="AnswerBox" FontSize="19" Margin="0,18,0,0" Padding="10,8"
                 BorderBrush="#C9D6E4" BorderThickness="1" FontFamily="Consolas"/>
        <TextBlock Name="FeedbackText" Text="" FontSize="12.5" Margin="0,12,0,0"
                   TextWrapping="Wrap" Foreground="#666666"/>
      </StackPanel>
      <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="26,4,26,20">
        <Button Name="BtnCancel" Content="取消" Background="#EDF2F8" Foreground="#1F4E79"/>
        <Button Name="BtnSubmit" Content="提交（回车）" Background="#1F4E79" Foreground="#FFFFFF"/>
      </StackPanel>
    </Grid>
  </Border>
</Window>
'@

    $win = [Windows.Markup.XamlReader]::Parse($xaml)
    $win.FindName('BtnCancel').Style = $win.Resources['FlatBtn']
    $win.FindName('BtnSubmit').Style = $win.Resources['FlatBtn']

    $txtPrompt = $win.FindName('PromptText')
    $txtPos = $win.FindName('PosText')
    $txtFeedback = $win.FindName('FeedbackText')
    $txtProgress = $win.FindName('ProgressText')
    $txtTitle = $win.FindName('TitleText')
    $box = $win.FindName('AnswerBox')
    $btnSubmit = $win.FindName('BtnSubmit')

    function Set-Feedback([string]$text, [string]$color) {
        $txtFeedback.Text = $text
        $txtFeedback.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString($color)
    }

    function Show-Word {
        $script:wrong = 0
        $script:revealed = $false
        $w = $words[$script:idx]
        if ($w.rev) { $txtTitle.Text = '拼写测试 · 复习词 · 全部拼对才能打卡' }
        else { $txtTitle.Text = '拼写测试 · 全部拼对才能打卡' }
        $roundText = ''
        if ($script:round -gt 1) { $roundText = ('第 {0} 轮 · ' -f $script:round) }
        $txtProgress.Text = ($roundText + ("第 {0} / {1} 个 · 一次通过 {2} 个" -f ($script:idx + 1), $total, $script:firstTry))
        if ($w.rev) { $txtPrompt.Text = ('【复习】' + $w.cn) } else { $txtPrompt.Text = $w.cn }
        $txtPos.Text = $w.pos
        $box.Text = ''
        $box.IsEnabled = $true
        $box.Focus() | Out-Null
        Set-Feedback '看中文，默写英文，回车提交' '#8C8C8C'
    }

    function Test-Answer {
        if ($script:busy) { return }
        if ($script:idx -ge $total) { return }
        $w = $words[$script:idx]
        $typed = $box.Text.Trim()
        if ($typed -eq '') { return }
        $target = $w.w.ToLower()
        if ($typed.ToLower() -eq $target) {
            if ($script:wrong -eq 0) { $script:firstTry++ }
            $script:busy = $true
            $box.IsEnabled = $false
            Set-Feedback ('✓ 正确：' + $w.w) '#2E7D4F'
            $script:next = New-Object Windows.Threading.DispatcherTimer
            $script:next.Interval = [TimeSpan]::FromMilliseconds(420)
            $script:next.Add_Tick({
                $script:next.Stop()
                $script:busy = $false
                $script:idx++
                if ($script:idx -ge $total) {
                    $script:passed = $true
                    $txtTitle.Text = '✓ 全部拼写正确'
                    if ($script:round -gt 1) {
                        $txtTitle.Text = ('✓ 全部拼写正确（第 {0} 轮通过）' -f $script:round)
                    }
                    $txtProgress.Text = ("{0} / {0} · 一次通过 {1} 个" -f $total, $script:firstTry)
                    $txtPrompt.Text = '已记录，正在打卡…'
                    $txtPos.Text = ''
                    $box.Visibility = 'Collapsed'
                    $box.IsEnabled = $false
                    $btnSubmit.IsEnabled = $false
                    Set-Feedback '' '#2E7D4F'
                    $script:done = New-Object Windows.Threading.DispatcherTimer
                    $script:done.Interval = [TimeSpan]::FromMilliseconds(1100)
                    $script:done.Add_Tick({ $script:done.Stop(); $script:win.Close() })
                    $script:done.Start()
                } else {
                    Show-Word
                }
            })
            $script:next.Start()
            return
        }

        $script:wrong++
        if ($script:wrong -eq 1) {
            Set-Feedback '✗ 不对，再试一次' '#C0392B'
        } elseif ($script:wrong -eq 2) {
            $hint = '首字母是 ' + $w.w.Substring(0, 1) + '，共 ' + $w.w.Length + ' 个字母'
            Set-Feedback ('✗ 再想想：' + $hint) '#C0392B'
        } else {
            # third miss: the round is void - close and let the user start over
            $script:busy = $true
            $script:void = $true
            $box.IsEnabled = $false
            $txtTitle.Text = '✗ 本轮作废'
            Set-Feedback ('3 次都没拼对，答案是 ' + $w.w + '。回到单词页点「今天已完成」重新开始。') '#C0392B'
            $script:endTimer = New-Object Windows.Threading.DispatcherTimer
            $script:endTimer.Interval = [TimeSpan]::FromMilliseconds(2200)
            $script:endTimer.Add_Tick({
                $script:endTimer.Stop()
                $script:win.Close()
            })
            $script:endTimer.Start()
            return
        }
        $box.SelectAll()
        $box.Focus() | Out-Null
    }

    $btnSubmit.Add_Click({ Test-Answer })
    $box.Add_KeyDown({
            param($s, $e)
            if ($e.Key -eq [Windows.Input.Key]::Enter) { Test-Answer; $e.Handled = $true }
        })
    $win.FindName('BtnCancel').Add_Click({ $script:passed = $false; $script:win.Close() })
    $win.Add_KeyDown({
            param($s, $e)
            if ($e.Key -eq [Windows.Input.Key]::Escape) { $script:passed = $false; $script:win.Close() }
        })
    $win.FindName('Header').Add_MouseLeftButtonDown({ try { $script:win.DragMove() } catch { } })
    $win.Add_Loaded({
            try { $script:win.Activate() } catch { }
            $script:box.Focus() | Out-Null
        })

    Show-Word

    if ($Test) {
        $win.WindowStartupLocation = 'Manual'
        $win.Left = -3000
        $win.Top = -3000
        $script:shot = New-Object Windows.Threading.DispatcherTimer
        $script:shot.Interval = [TimeSpan]::FromMilliseconds(700)
        $script:shot.Add_Tick({
            try {
                $rtb = New-Object Windows.Media.Imaging.RenderTargetBitmap -ArgumentList @(
                    [int]($script:win.ActualWidth * 2), [int]($script:win.ActualHeight * 2), 192, 192,
                    [Windows.Media.PixelFormats]::Pbgra32)
                $rtb.Render($script:win)
                $enc = New-Object Windows.Media.Imaging.PngBitmapEncoder
                $enc.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($rtb))
                $fs = [System.IO.File]::Create((Join-Path $script:root 'preview-spell.png'))
                $enc.Save($fs)
                $fs.Close()
                Write-Log '拼写测试截图完成'
            } catch {
                Write-Log ('拼写测试截图失败: ' + $_.Exception.Message)
            }
            $script:shot.Stop()
            $script:win.Close()
        })
        $script:shot.Start()
        $win.ShowDialog() | Out-Null
        exit 0
    }

    $win.ShowDialog() | Out-Null
    if ($script:passed) { Write-Log ('拼写测试通过 第' + $Day + '天'); exit 0 }
    if ($script:void) { Write-Log ('拼写测试本轮作废 第' + $Day + '天'); exit 1 }
    Write-Log ('拼写测试取消 第' + $Day + '天')
    exit 2
}
catch {
    Write-Log ('拼写测试出错: ' + $_.Exception.Message)
    exit 3
}
