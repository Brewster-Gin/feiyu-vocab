# Shift the whole 40-day plan to a new start date.
# Rewrites words.json day dates + the workbook's date cells.
function Fix-PlanDates {
    param(
        [Parameter(Mandatory = $true)][string]$Dir,
        [Parameter(Mandatory = $true)][datetime]$Start
    )

    $epoch = [datetime]'1899-12-30'
    $exam = $Start.AddDays(40)

    $jsonPath = Join-Path $Dir 'words.json'
    $plan = Get-Content $jsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $i = 0
    foreach ($d in $plan.days) {
        $d.date = $Start.AddDays($i).ToString('yyyy-MM-dd')
        $i++
    }
    $plan.start = $Start.ToString('yyyy-MM-dd')
    $plan.exam = $exam.ToString('yyyy-MM-dd')
    $json = $plan | ConvertTo-Json -Depth 8
    [System.IO.File]::WriteAllText($jsonPath, $json, (New-Object System.Text.UTF8Encoding($false)))

    $xlsx = Get-ChildItem -Path $Dir -Filter '*.xlsx' -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $xlsx) { throw 'workbook not found' }

    Add-Type -AssemblyName System.IO.Compression.FileSystem, System.IO.Compression -ErrorAction SilentlyContinue
    $tmp = Join-Path $env:TEMP ('cet6_shift_' + [guid]::NewGuid().ToString('N') + '.xlsx')
    Copy-Item $xlsx.FullName $tmp -Force
    try {
        $zip = [System.IO.Compression.ZipFile]::Open($tmp, [System.IO.Compression.ZipArchiveMode]::Update)
        try {
            $jobs = @{}
            $jobs['xl/worksheets/sheet1.xml'] = @{}
            for ($r = 2; $r -le 41; $r++) {
                $jobs['xl/worksheets/sheet1.xml']['B' + $r] = [int]($Start.AddDays($r - 2) - $epoch).TotalDays
            }
            $jobs['xl/worksheets/sheet2.xml'] = @{}
            for ($r = 2; $r -le 401; $r++) {
                $dayIndex = [math]::Floor(($r - 2) / 10)
                $jobs['xl/worksheets/sheet2.xml']['C' + $r] = [int]($Start.AddDays($dayIndex) - $epoch).TotalDays
            }
            $jobs['xl/worksheets/sheet3.xml'] = @{
                'B2' = [int]($Start - $epoch).TotalDays
                'B3' = [int]($exam - $epoch).TotalDays
            }

            foreach ($name in $jobs.Keys) {
                $entry = $zip.GetEntry($name)
                if (-not $entry) { continue }
                $sr = New-Object System.IO.StreamReader($entry.Open())
                $xml = $sr.ReadToEnd()
                $sr.Close()
                $entry.Delete()

                foreach ($ref in $jobs[$name].Keys) {
                    $val = $jobs[$name][$ref]
                    $pat = '<c r="' + $ref + '"([^>]*)><v>[^<]*</v></c>'
                    $rep = '<c r="' + $ref + '"$1><v>' + $val + '</v></c>'
                    $rx = [regex]::new($pat)
                    if ($rx.IsMatch($xml)) {
                        $xml = $rx.Replace($xml, $rep, 1)
                    }
                }

                $newEntry = $zip.CreateEntry($name)
                $sw = New-Object System.IO.StreamWriter($newEntry.Open(), (New-Object System.Text.UTF8Encoding($false)))
                $sw.Write($xml)
                $sw.Close()
            }
        } finally { $zip.Dispose() }
        Copy-Item $tmp $xlsx.FullName -Force
    } finally {
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    }

    return $exam
}
