# 銘柄数の多いユニバース向けに、日ごとの売買代金上位N銘柄だけを作る軽量版。
# Build-Features.ps1 と同じ定義（終値×出来高の turnoverMaWindow 日平均、各銘柄の最初の rankingWarmupDays 日は対象外）で、
# 全銘柄の全行をメモリに載せず、銘柄ファイルを1つずつ読んで日ごとの上位Nだけを保持する。
# 使い方: $env:KABU_CONFIG = "ps/config.prime.json"; powershell -File ps\Build-TurnoverRanking.ps1

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"

$config = Get-Config
$topN = [int]$config.universe.topNPopular
$maWindow = [int]$config.features.turnoverMaWindow
$warmup = 75
if ($config.features.rankingWarmupDays) { $warmup = [int]$config.features.rankingWarmupDays }
$universe = Import-Csv -Path (Resolve-ProjectPath $config.universe.csv) -Encoding UTF8
# universe.changesCsv があれば、各日時点で指数に入っていた銘柄だけを候補にする
$membership = Import-IndexMembership -Config $config
if ($membership) { Write-Host "using point-in-time membership from $($config.universe.changesCsv)" }
$tickerDir = Split-Path (Resolve-ProjectPath "$($config.data.tickerDir)/placeholder") -Parent

# date -> 売買代金平均の降順に並んだ上位N件
$top = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[object]]'
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$used = 0; $missing = 0; $k = 0

foreach ($u in $universe) {
    $k++
    $path = Join-Path $tickerDir "$($u.code).csv"
    if (-not (Test-Path $path)) { $missing++; continue }
    $rows = @(Import-Csv -Path $path -Encoding UTF8)
    $n = $rows.Count
    if ($n -lt ($warmup + 10)) { continue }
    $used++

    # 値が不正な行は移動平均の窓をリセットし、不連続(再上場・分割の調整漏れ等)の後は
    # 新しい銘柄として warmup からやり直す（判定基準は Common.ps1 の Set-InvalidPriceRows と同じ）
    $win = New-Object 'System.Collections.Generic.Queue[double]'
    $sum = 0.0
    $segValid = 0
    $lastValid = -1
    for ($i = 0; $i -lt $n; $i++) {
        $o = [double]$rows[$i].open
        $c = [double]$rows[$i].close
        $v = [double]$rows[$i].volume
        $valid = Test-PriceRowValid -Open $o -Close $c -RawClose ([double]$rows[$i].raw_close) -Volume $v
        $isBreak = $false
        if ($valid -and $lastValid -ge 0) {
            if (([datetime]$rows[$i].date - [datetime]$rows[$lastValid].date).TotalDays -gt $script:MaxDateGapDays) { $isBreak = $true }
            elseif ($lastValid -eq ($i - 1)) {
                $ratio = $o / [double]$rows[$lastValid].close
                if ($ratio -gt $script:MaxOvernightRatio -or $ratio -lt $script:MinOvernightRatio) { $isBreak = $true }
            }
        }
        if (-not $valid -or $isBreak) {
            $win.Clear(); $sum = 0.0
            if ($isBreak) { $segValid = 0; $lastValid = -1 }
            continue
        }
        $lastValid = $i
        $segValid++
        $t = $c * $v
        $win.Enqueue($t); $sum += $t
        if ($win.Count -gt $maWindow) { $sum -= $win.Dequeue() }
        if ($segValid -le $warmup -or $win.Count -lt $maWindow) { continue }
        $tm = $sum / $maWindow
        $turnoverToday = $t
        $date = $rows[$i].date
        if ($membership -and -not (Test-IndexMember -Membership $membership -Code $u.code -Date $date)) { continue }
        $list = $null
        if (-not $top.TryGetValue($date, [ref]$list)) {
            $list = New-Object System.Collections.Generic.List[object]
            $top[$date] = $list
        }
        # 満杯で最下位以下なら入らない（大半の銘柄はここで弾かれる）
        if ($list.Count -ge $topN -and $tm -le $list[$list.Count - 1].turnover_ma) { continue }
        $entry = [PSCustomObject]@{ code = $u.code; close = $c; turnover = $turnoverToday; turnover_ma = $tm }
        $pos = $list.Count
        for ($j = 0; $j -lt $list.Count; $j++) {
            if ($tm -gt $list[$j].turnover_ma) { $pos = $j; break }
        }
        $list.Insert($pos, $entry)
        if ($list.Count -gt $topN) { $list.RemoveAt($list.Count - 1) }
    }
    if ($k % 200 -eq 0) { Write-Host ("  {0}/{1} tickers, elapsed {2}" -f $k, $universe.Count, $sw.Elapsed) }
}

$nameByCode = @{}; $sectorByCode = @{}
foreach ($u in $universe) { $nameByCode[$u.code] = $u.name; $sectorByCode[$u.code] = $u.sector }

$out = New-Object System.Collections.Generic.List[object]
foreach ($date in ($top.Keys | Sort-Object)) {
    $list = $top[$date]
    for ($r = 0; $r -lt $list.Count; $r++) {
        $e = $list[$r]
        $out.Add([PSCustomObject]@{ date = $date; rank = $r + 1; code = $e.code; name = $nameByCode[$e.code]; sector = $sectorByCode[$e.code]; close = $e.close; turnover = $e.turnover; turnover_ma = $e.turnover_ma })
    }
}
$outPath = Resolve-ProjectPath $config.data.dailyRankingCsv
$out | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8
Write-Host ("saved {0} rows ({1} days) to {2}. tickers used {3}, missing files {4}, elapsed {5}" -f $out.Count, $top.Count, $outPath, $used, $missing, $sw.Elapsed)
