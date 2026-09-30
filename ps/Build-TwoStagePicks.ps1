# B3（低ボラ条件なし）と B2（B3L）の上位10銘柄から、前向き検証用の2系統の date,rank,code を作る。
#   b3_b3l_m0.csv : M0 = B3 と B2 の上位10ずつを半分ずつ（重複する銘柄は重み2、その他は1。weight 列つき）
#   t2_on20.csv   : T2 = 上の約20銘柄を「直近20営業日の夜間リターン平均」で並べ替えた上位10
# 夜間リターン = 調整後の始値 ÷ 前営業日の調整後の終値 − 1。20個が揃わない銘柄（|夜間リターン| が 40% 以上の日を含む窓を含む）は最下位。
# 同値はコード順。判断日 t の引けまでのデータだけを使う。
# 先に Export-B3LPicks.ps1（B3L）と Export-B3LPicks.ps1 -Column picks_B3 -OutDir data/processed/b3（B3）を実行しておく。
# 使い方: powershell -File ps\Build-TwoStagePicks.ps1
param(
    [string]$B3LCsv   = "data/processed/b3l/daily_picks.csv",
    [string]$B3Csv    = "data/processed/b3/daily_picks.csv",
    [string]$TickerDir = "data/raw/stocks/prime_since2000",
    [string]$OutDir   = "data/processed/combo",
    # この日以降の判断日だけ作る（夜間リターンの20営業日は、これより前のデータも読む）
    [string]$FromDate = "2025-12-01"
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"
Set-Location (Get-ProjectRoot)

function Import-RankedCsv {
    # date -> rank順に並べた code の配列
    param([string]$Path)
    $full = Resolve-ProjectPath $Path
    if (-not (Test-Path $full)) { throw "$Path がない" }
    $map = @{}
    foreach ($r in (Import-Csv -Path $full -Encoding UTF8)) {
        if (-not $map.ContainsKey($r.date)) { $map[$r.date] = New-Object System.Collections.Generic.List[object] }
        $map[$r.date].Add([PSCustomObject]@{ rank = [int]$r.rank; code = [string]$r.code })
    }
    $out = @{}
    foreach ($d in $map.Keys) { $out[$d] = @($map[$d] | Sort-Object rank | Select-Object -ExpandProperty code) }
    return $out
}

$b3l = Import-RankedCsv $B3LCsv
$b3  = Import-RankedCsv $B3Csv
$dates = @($b3l.Keys | Where-Object { $b3.ContainsKey($_) -and [string]::CompareOrdinal($_, $FromDate) -ge 0 -and $b3[$_].Count -ge 10 -and $b3l[$_].Count -ge 10 } | Sort-Object)
if ($dates.Count -eq 0) { throw "B3 と B3L の上位10がそろった日がない" }

# 必要な銘柄の夜間リターン20日平均（判断日ごと）を作る
$need = New-Object System.Collections.Generic.HashSet[string]
foreach ($d in $dates) { foreach ($c in $b3[$d]) { [void]$need.Add($c) }; foreach ($c in $b3l[$d]) { [void]$need.Add($c) } }
$dateSet = New-Object System.Collections.Generic.HashSet[string]
foreach ($d in $dates) { [void]$dateSet.Add($d) }
$tickerFull = Split-Path (Resolve-ProjectPath "$TickerDir/placeholder") -Parent
$on20 = @{}   # code -> (date -> 20日平均)。無い日は NaN 扱い
foreach ($code in $need) {
    $path = Join-Path $tickerFull "$code.csv"
    $byDate = New-Object 'System.Collections.Generic.Dictionary[string,double]'
    if (Test-Path $path) {
        $rows = Import-Csv -Path $path -Encoding UTF8
        $n = $rows.Count
        $on = [double[]]::new($n)
        $prevC = [double]::NaN
        for ($i = 0; $i -lt $n; $i++) {
            $o = [double]$rows[$i].open; $c = [double]$rows[$i].close
            $v = [double]::NaN
            if ($o -gt 0 -and $c -gt 0) {
                if (-not [double]::IsNaN($prevC)) {
                    $x = $o / $prevC - 1.0
                    if ([Math]::Abs($x) -lt 0.4) { $v = $x }
                }
                $prevC = $c
            } else { $prevC = [double]::NaN }
            $on[$i] = $v
        }
        for ($i = 19; $i -lt $n; $i++) {
            $d = [string]$rows[$i].date
            if (-not $dateSet.Contains($d)) { continue }
            $s = 0.0; $ok = $true
            for ($k = $i - 19; $k -le $i; $k++) { if ([double]::IsNaN($on[$k])) { $ok = $false; break }; $s += $on[$k] }
            if ($ok) { $byDate[$d] = $s / 20.0 }
        }
    }
    $on20[$code] = $byDate
}

$m0Rows = New-Object System.Collections.Generic.List[object]
$t2Rows = New-Object System.Collections.Generic.List[object]
foreach ($d in $dates) {
    # 候補: B3 の上位10 → B2 の上位10（重複は1つ）。M0 の重みは B3 と B2 に入った回数
    $codes = New-Object System.Collections.Generic.List[string]
    $weight = @{}
    foreach ($c in ($b3[$d] | Select-Object -First 10)) { if (-not $weight.ContainsKey($c)) { $codes.Add($c); $weight[$c] = 0.0 }; $weight[$c] += 1.0 }
    foreach ($c in ($b3l[$d] | Select-Object -First 10)) { if (-not $weight.ContainsKey($c)) { $codes.Add($c); $weight[$c] = 0.0 }; $weight[$c] += 1.0 }
    $k = 1
    foreach ($c in $codes) { $m0Rows.Add([PSCustomObject]@{ date = $d; rank = $k; code = $c; weight = $weight[$c] }); $k++ }

    $items = New-Object System.Collections.Generic.List[object]
    foreach ($c in $codes) {
        $score = [double]::NegativeInfinity
        $val = 0.0
        if ($on20.ContainsKey($c) -and $on20[$c].TryGetValue($d, [ref]$val)) { $score = $val }
        $items.Add([PSCustomObject]@{ code = $c; score = $score })
    }
    $top = @($items | Sort-Object @{ Expression = "score"; Descending = $true }, @{ Expression = "code"; Descending = $false } | Select-Object -First 10)
    $k = 1
    foreach ($t in $top) { $t2Rows.Add([PSCustomObject]@{ date = $d; rank = $k; code = $t.code }); $k++ }
}

$outFull = Split-Path (Resolve-ProjectPath "$OutDir/placeholder") -Parent
New-Item -ItemType Directory -Force $outFull | Out-Null
$m0Rows | Export-CsvNoBom -Path (Join-Path $outFull "b3_b3l_m0.csv")
$t2Rows | Export-CsvNoBom -Path (Join-Path $outFull "t2_on20.csv")
Write-Host ("b3_b3l_m0.csv {0,6} rows / {1,4} days  {2} .. {3}" -f $m0Rows.Count, $dates.Count, $dates[0], $dates[-1])
Write-Host ("t2_on20.csv   {0,6} rows / {1,4} days  {2} .. {3}" -f $t2Rows.Count, $dates.Count, $dates[0], $dates[-1])
