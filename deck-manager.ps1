# 词库管理：列出 / 切换 / 导入
#   双击 词库管理.cmd 运行
param(
    [switch]$ListOnly,
    [string]$SwitchTo = '',
    [string]$ImportFile = ''
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$deckDir = Join-Path $root '词库'
$buildPy = Join-Path $root '工具\build_plan.py'
if (-not (Test-Path $deckDir)) { New-Item -ItemType Directory -Path $deckDir -Force | Out-Null }

function Find-Python {
    $dsh = Join-Path $env:USERPROFILE '.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe'
    if (Test-Path $dsh) { return $dsh }
    $c = Get-Command python -ErrorAction SilentlyContinue
    if ($c) { return $c.Source }
    return $null
}

function Get-Decks {
    Get-ChildItem $deckDir -Filter '*.csv' -ErrorAction SilentlyContinue | Sort-Object Name
}

function Show-Decks {
    $decks = @(Get-Decks)
    if ($decks.Count -eq 0) { Write-Host '   (词库目录是空的)' -ForegroundColor DarkGray; return }
    for ($i = 0; $i -lt $decks.Count; $i++) {
        $lines = @(Get-Content $decks[$i].FullName -Encoding UTF8)
        $count = [Math]::Max(0, $lines.Count - 1)
        $active = ''
        if ($decks[$i].Name -eq $script:activeDeck) { $active = '  ← 当前' }
        Write-Host ("   [{0}] {1}   （{2} 词）{3}" -f ($i + 1), $decks[$i].BaseName, $count, $active) -ForegroundColor Cyan
    }
}

function Get-ActiveDeck {
    $cfg = Join-Path $root 'config.json'
    if (Test-Path $cfg) {
        try { return (Get-Content $cfg -Raw -Encoding UTF8 | ConvertFrom-Json).deck } catch { }
    }
    return ''
}

function Save-ActiveDeck([string]$name) {
    $cfg = Join-Path $root 'config.json'
    $obj = @{}
    if (Test-Path $cfg) {
        try {
            $old = Get-Content $cfg -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($p in $old.PSObject.Properties) { $obj[$p.Name] = $p.Value }
        } catch { }
    }
    $obj['deck'] = $name
    ($obj | ConvertTo-Json) | Set-Content $cfg -Encoding UTF8
}

function Invoke-Build([string]$deckPath) {
    $py = Find-Python
    if (-not $py) { Write-Host '[!!] 找不到 Python，无法生成计划。' -ForegroundColor Yellow; return $false }
    $today = (Get-Date).ToString('yyyy-MM-dd')
    Write-Host '   正在生成计划（约 2 秒）...' -ForegroundColor DarkGray
    & $py $buildPy --deck $deckPath --per-day 10 --intervals 1,4,15 --start $today
    if ($LASTEXITCODE -ne 0) { Write-Host '[!!] 生成失败。' -ForegroundColor Yellow; return $false }
    return $true
}

function Switch-Deck([string]$deckPath) {
    $name = [System.IO.Path]::GetFileNameWithoutExtension($deckPath)
    Write-Host ''
    Write-Host ("切换到词库：$name") -ForegroundColor Cyan
    Write-Host '注意：会重建打卡表和 words.json，之前的打卡记录将对不上（学习进度会归零）。' -ForegroundColor Yellow
    $ok = Read-Host '确定切换吗？[y/N]'
    if ($ok -notmatch '^(y|Y)') { Write-Host '已取消。' -ForegroundColor DarkGray; return }
    if (Invoke-Build $deckPath) {
        Save-ActiveDeck $name
        $script:activeDeck = $name
        Write-Host ("[OK] 已切换到「$name」，共 " + ([Math]::Max(0, (Get-Content $deckPath -Encoding UTF8).Count - 1)) + ' 个词。') -ForegroundColor Green
        Write-Host '     下次弹窗就按新词库出题（当天已打卡的话，明天生效）。'
    }
}

function Import-Deck([string]$file) {
    if (-not (Test-Path $file)) { Write-Host ('[!!] 找不到文件: ' + $file) -ForegroundColor Yellow; return }
    $ext = [System.IO.Path]::GetExtension($file).ToLower()
    $rows = @()

    if ($ext -eq '.csv') {
        $all = @(Import-Csv -Path $file -Encoding UTF8)
        if ($all.Count -eq 0) { Write-Host '[!!] CSV 是空的。' -ForegroundColor Yellow; return }
        $map = @{}
        foreach ($k in $all[0].PSObject.Properties.Name) { $map[$k.Trim().ToLower()] = $k }
        $colW = @('word', '单词', '英文', '词汇') | Where-Object { $map.ContainsKey($_) } | Select-Object -First 1
        $colM = @('meaning', '释义', '中文', '中文意思', '翻译') | Where-Object { $map.ContainsKey($_) } | Select-Object -First 1
        if (-not $colW) { Write-Host '[!!] 认不出「单词」列（支持：word/单词/英文）。' -ForegroundColor Yellow; return }
        if (-not $colM) { Write-Host '[!!] 认不出「释义」列（支持：meaning/释义/中文）。' -ForegroundColor Yellow; return }
        $colP = @('pos', '词性') | Where-Object { $map.ContainsKey($_) } | Select-Object -First 1
        $colH = @('phrase', '搭配', '例句') | Where-Object { $map.ContainsKey($_) } | Select-Object -First 1
        $colI = @('ipa', '音标') | Where-Object { $map.ContainsKey($_) } | Select-Object -First 1
        foreach ($r in $all) {
            $w = $r.$($map[$colW]); $m = $r.$($map[$colM])
            if (-not $w -or -not $m) { continue }
            $rows += [pscustomobject]@{
                word = $w.Trim(); pos = $(if ($colP) { $r.$($map[$colP]) } else { '' })
                meaning = $m.Trim(); phrase = $(if ($colH) { $r.$($map[$colH]) } else { '' })
                ipa = $(if ($colI) { $r.$($map[$colI]) } else { '' })
            }
        }
    } elseif ($ext -eq '.txt') {
        foreach ($line in (Get-Content $file -Encoding UTF8)) {
            $t = $line.Trim()
            if (-not $t) { continue }
            $t = $t -replace '^\d+[.、) ]+', ''
            $w = ''; $ipa = ''; $m = ''
            if ($t -match '^([A-Za-z][A-Za-z''\-. ]*?)\s+([/\[]{1}[^/\]]{1,40}[/\]]{1})\s*(.+)$') {
                $w = $matches[1]; $ipa = $matches[2]; $m = $matches[3]
            }
            elseif ($t -match '^([A-Za-z][A-Za-z''\-.]*)\s+(.*[\u4e00-\u9fa5].*)$') {
                $w = $matches[1]; $m = $matches[2]
            }
            else {
                $parts = $t -split "`t"
                if ($parts.Count -lt 2) { $parts = $t -split '\s{2,}' }
                if ($parts.Count -lt 2) { $parts = $t -split ',' }
                if ($parts.Count -lt 2) { continue }
                $w = $parts[0].Trim(); $m = $parts[1].Trim()
            }
            if ($w -and $m) {
                $rows += [pscustomobject]@{ word = $w.Trim(); pos = ''; meaning = $m.Trim(); phrase = ''; ipa = $ipa }
            }
        }
    } else {
        Write-Host '[!!] 只支持 .csv 和 .txt' -ForegroundColor Yellow; return
    }

    if ($rows.Count -eq 0) { Write-Host '[!!] 没解析出任何「单词 + 释义」配对。' -ForegroundColor Yellow; return }

    $name = [System.IO.Path]::GetFileNameWithoutExtension($file)
    $target = Join-Path $deckDir ($name + '.csv')
    $rows | Select-Object word, pos, meaning, phrase, ipa |
        Export-Csv -Path $target -NoTypeInformation -Encoding UTF8
    Write-Host ("[OK] 已导入 " + $rows.Count + ' 个词 → 词库\' + $name + '.csv') -ForegroundColor Green
}

$script:activeDeck = Get-ActiveDeck

if ($SwitchTo) { Switch-Deck (Join-Path $deckDir ($SwitchTo + '.csv')); exit 0 }
if ($ImportFile) { Import-Deck $ImportFile; exit 0 }
if ($ListOnly) { Show-Decks; exit 0 }

Write-Host ''
Write-Host '==== 词库管理 ====' -ForegroundColor Cyan
Write-Host ("当前词库: " + $(if ($script:activeDeck) { $script:activeDeck } else { '（默认 400 词）' }))
Write-Host ''
Write-Host '可用词库:' -ForegroundColor White
Show-Decks
Write-Host ''
Write-Host '  输入序号 = 切换词库'
Write-Host '  输入 i   = 导入新词表（csv / txt）'
Write-Host '  回车     = 退出'
Write-Host ''

while ($true) {
    $ans = Read-Host '请选择'
    if (-not $ans) { break }
    if ($ans -eq 'i') {
        $f = Read-Host '把文件拖到这里（或输入完整路径）'
        $f = $f.Trim('"').Trim()
        Import-Deck $f
        continue
    }
    $n = 0
    if ([int]::TryParse($ans, [ref]$n)) {
        $decks = @(Get-Decks)
        if ($n -ge 1 -and $n -le $decks.Count) {
            Switch-Deck $decks[$n - 1].FullName
        } else {
            Write-Host '序号超出范围。' -ForegroundColor Yellow
        }
    }
}
