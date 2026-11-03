# Register the CET-6 reminder as a Scheduled Task (needs an elevated shell / one UAC click).
# Triggers: logon (+40s) / workstation unlock / every day 20:30
#   powershell -NoProfile -ExecutionPolicy Bypass -File install-task.ps1
param([string]$Name = 'CET6-Word-Reminder')

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$rem = Join-Path $root 'reminder.ps1'
$vbs = Join-Path $root 'launch.vbs'
$xmlPath = Join-Path $root 'task.xml'
$userId = "$env:USERDOMAIN\$env:USERNAME"

if (-not (Test-Path $rem)) { throw "reminder.ps1 not found" }
if (-not (Test-Path $vbs)) { throw "launch.vbs not found" }

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
Write-Host ("Running elevated: " + $isAdmin)
if (-not $isAdmin) { Write-Host 'WARNING: not elevated, registration will most likely fail.' }

# ---------- approach 1: XML with all three triggers ----------
$xml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Author>$userId</Author>
    <Description>CET-6 daily word reminder popup</Description>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
      <UserId>$userId</UserId>
      <Delay>PT40S</Delay>
    </LogonTrigger>
    <SessionStateChangeTrigger>
      <Enabled>true</Enabled>
      <UserId>$userId</UserId>
      <StateChange>SessionUnlock</StateChange>
    </SessionStateChangeTrigger>
    <CalendarTrigger>
      <StartBoundary>2026-11-02T20:30:00</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByDay>
        <DaysInterval>1</DaysInterval>
      </ScheduleByDay>
    </CalendarTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>$userId</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT2H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>C:\Windows\System32\wscript.exe</Command>
      <Arguments>"$vbs"</Arguments>
      <WorkingDirectory>$root</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"@

[System.IO.File]::WriteAllText($xmlPath, $xml, (New-Object System.Text.UnicodeEncoding($false, $true)))
[xml]$check = Get-Content $xmlPath -Raw -Encoding Unicode
Write-Host ("XML parsed OK, triggers: " + $check.Task.Triggers.ChildNodes.Count)

$ok = $false
$ErrorActionPreference = 'Continue'
$out = (& schtasks.exe /create /tn $Name /xml $xmlPath /f 2>&1 | Out-String)
$code = $LASTEXITCODE
$ErrorActionPreference = 'Stop'
Write-Host $out.Trim()
if ($code -eq 0 -and $out -notmatch 'ERROR') { $ok = $true }

# ---------- approach 2: cmdlet fallback (logon + daily, no unlock) ----------
if (-not $ok) {
    Write-Host ''
    Write-Host 'XML import failed, falling back to Register-ScheduledTask ...'
    try {
        $action = New-ScheduledTaskAction -Execute 'C:\Windows\System32\wscript.exe' `
            -Argument ('"' + $vbs + '"') `
            -WorkingDirectory $root
        $t1 = New-ScheduledTaskTrigger -AtLogOn -User $userId
        $t2 = New-ScheduledTaskTrigger -Daily -At '20:30'
        $set = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
            -MultipleInstances IgnoreNew -StartWhenAvailable
        Register-ScheduledTask -TaskName $Name -Action $action -Trigger @($t1, $t2) -Settings $set `
            -Description 'CET-6 daily word reminder popup' -Force | Out-Null
        $ok = $true
        Write-Host 'Registered with cmdlets (logon + daily 20:30).'
    } catch {
        Write-Host ('Register-ScheduledTask failed: ' + $_.Exception.Message)
    }
}

if (-not $ok) {
    Write-Host ''
    Write-Host 'FAILED. Right-click 安装-任务计划.cmd and choose "Run as administrator".'
    exit 1
}

# remove the old startup-folder entry so nothing fires twice
$startup = [Environment]::GetFolderPath('Startup')
$old = Join-Path $startup 'CET6-Word-Reminder.vbs'
if (Test-Path $old) { Remove-Item $old -Force; Write-Host "Removed old startup entry: $old" }

Write-Host ''
Write-Host 'Done. Check it with:  schtasks /query /tn "CET6-Word-Reminder" /v /fo LIST'
Write-Host ''
Write-Host 'Running it once now so you can see the popup ...'
$ErrorActionPreference = 'Continue'
& schtasks.exe /run /tn $Name 2>&1 | Out-String | Write-Host
$ErrorActionPreference = 'Stop'
