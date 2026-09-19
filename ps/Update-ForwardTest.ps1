# 未知データ(モデルを固めた後に出てきたデータ)での成績を、追記専用の台帳に残す。
#
# 毎日シミュレーションをやり直すと、分割・配当で過去の調整済み価格が変わったときに
# 昔の数字まで動いてしまい、「その時どうだったか」の記録にならない。
# そこで一度書いた行は二度と書き換えず、まだ無い日付だけを足す。
#
#   reports/forward_test/daily.csv   買った日ごと・買い方ごとの成績(1行=1日1系統)
#   reports/forward_test/trades.csv  その内訳(1行=1建玉)
#   FORWARD_TEST_RESULTS.md          上の台帳から作る読み物
#
# 使い方: powershell -File ps\Update-ForwardTest.ps1
param(
    # この日に買った分から先が「未知データ」。既定の2026-09-17は、配布時点の
    # カレンダーが持っていた最後の建玉(9/16買い)の次の営業日。
    [string]$StartDate  = "2026-09-17",
    [string]$ReportRoot = "reports/b3l_2026",
    [string]$OutDir     = "reports/forward_test",
    [string]$SummaryMd  = "FORWARD_TEST_RESULTS.md",
    [int]$CostBps       = 3,
    [int]$Account       = 500000,
    # 既に書いた行も作り直す(台帳を作り直したいときだけ)
    [switch]$Rebuild
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"
Set-Location (Get-ProjectRoot)

$variants = @(
    @{ key = "b2_3";    dir = "top3";    scenario = "top3_compound";    label = "B2 3銘柄" }
    @{ key = "b2_5";    dir = "top5";    scenario = "top5_compound";    label = "B2 5銘柄" }
    @{ key = "b2_10";   dir = "top10";   scenario = "top10_compound";   label = "B2 10銘柄" }
    @{ key = "nk10";    dir = "nk10";    scenario = "nk10_compound";    label = "日経225 上位10" }
    @{ key = "prime10"; dir = "prime10"; scenario = "prime10_compound"; label = "プライム 上位10" }
    @{ key = "mix55";   dir = "mix55";   scenario = "mix55_compound";   label = "日経225上位5＋B2上位5" }
)

$names = New-Object 'System.Collections.Generic.Dictionary[string,string]'
foreach ($u in (Import-Csv -Path (Resolve-ProjectPath "data/raw/universe/prime.csv") -Encoding UTF8)) { $names[[string]$u.code] = [string]$u.name }

$outDirFull = Split-Path (Resolve-ProjectPath "$OutDir/placeholder") -Parent
New-Item -ItemType Directory -Force $outDirFull | Out-Null
$dailyPath  = Join-Path $outDirFull "daily.csv"
$tradesPath = Join-Path $outDirFull "trades.csv"

# --- 既にある台帳を読む(キーは 買い日+系統 / 買い日+系統+銘柄) ---
$haveDay   = New-Object System.Collections.Generic.HashSet[string]
$haveTrade = New-Object System.Collections.Generic.HashSet[string]
$dailyRows = New-Object System.Collections.Generic.List[object]
$tradeRows = New-Object System.Collections.Generic.List[object]
if (-not $Rebuild) {
    if (Test-Path $dailyPath) {
        foreach ($r in (Import-Csv -Path $dailyPath -Encoding UTF8)) {
            $dailyRows.Add($r)
            [void]$haveDay.Add("$($r.buy_date)|$($r.variant)")
        }
    }
    if (Test-Path $tradesPath) {
        foreach ($r in (Import-Csv -Path $tradesPath -Encoding UTF8)) {
            $tradeRows.Add($r)
            [void]$haveTrade.Add("$($r.buy_date)|$($r.variant)|$($r.code)")
        }
    }
}
$before = $dailyRows.Count
$recordedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

foreach ($v in $variants) {
    $tag       = "$($v.scenario)_cost$CostBps"
    $tradesCsv = Resolve-ProjectPath "$ReportRoot/$($v.dir)/trades_$tag.csv"
    $dailyCsv  = Resolve-ProjectPath "$ReportRoot/$($v.dir)/daily_$tag.csv"
    if (-not (Test-Path $tradesCsv) -or -not (Test-Path $dailyCsv)) {
        Write-Warning "skip $($v.key): シミュレーション結果がない"
        continue
    }

    # PowerShell の @{} はキーの型でつまずくことがあるので、型を決めた辞書を使う
    $equity = New-Object 'System.Collections.Generic.Dictionary[string,double]'
    foreach ($r in (Import-Csv -Path $dailyCsv -Encoding UTF8)) { $equity[[string]$r.date] = [double]$r.equity }

    # 買った日ごとにまとめる
    $rowsAll = @(Import-Csv -Path $tradesCsv -Encoding UTF8 | Where-Object { [string]::Compare([string]$_.buy_date, $StartDate) -ge 0 })

    foreach ($g in ($rowsAll | Group-Object buy_date | Sort-Object Name)) {
        $d  = [string]$g.Name
        $ts = @($g.Group)
        $sellDate = [string]$ts[0].sell_date
        $key = "$d|$($v.key)"

        foreach ($t in ($ts | Sort-Object { -[double]$_.pnl })) {
            $code = [string]$t.code
            $tk = "$d|$($v.key)|$code"
            if ($haveTrade.Contains($tk)) { continue }
            $nm = $code
            if ($names.ContainsKey($code)) { $nm = $names[$code] }
            $tradeRows.Add([PSCustomObject]@{
                buy_date    = $d
                sell_date   = [string]$t.sell_date
                variant     = $v.key
                code        = $code
                name        = $nm
                shares      = [int]$t.shares
                buy_price   = [Math]::Round([double]$t.buy_price, 1)
                sell_price  = [Math]::Round([double]$t.sell_price_raw, 1)
                buy_amount  = [int][Math]::Round([double]$t.buy_amount)
                sell_amount = [int][Math]::Round([double]$t.proceeds_div_adjusted)
                fee         = [Math]::Round([double]$t.fee, 2)
                pnl         = [int][Math]::Round([double]$t.pnl)
                recorded_at = $recordedAt
            })
            [void]$haveTrade.Add($tk)
        }

        if ($haveDay.Contains($key)) { continue }
        $buyAmt = 0.0; $sellAmt = 0.0; $fee = 0.0; $pnl = 0.0
        foreach ($t in $ts) {
            $buyAmt  += [double]$t.buy_amount
            $sellAmt += [double]$t.proceeds_div_adjusted
            $fee     += [double]$t.fee
            $pnl     += [double]$t.pnl
        }
        $eqAfter = [double]::NaN
        if ($equity.ContainsKey($sellDate)) { $eqAfter = $equity[$sellDate] }
        $dailyRows.Add([PSCustomObject]@{
            buy_date     = $d
            sell_date    = $sellDate
            variant      = $v.key
            label        = $v.label
            stocks       = $ts.Count
            buy_amount   = [int][Math]::Round($buyAmt)
            sell_amount  = [int][Math]::Round($sellAmt)
            fee          = [Math]::Round($fee, 2)
            pnl          = [int][Math]::Round($pnl)
            return_pct   = [Math]::Round($pnl / $buyAmt * 100, 4)
            equity_after = [int][Math]::Round($eqAfter)
            recorded_at  = $recordedAt
        })
        [void]$haveDay.Add($key)
    }
}

$dailySorted  = @($dailyRows | Sort-Object buy_date, variant)
$tradesSorted = @($tradeRows | Sort-Object buy_date, variant, @{ Expression = { -[double]$_.pnl } })
$dailySorted  | Export-Csv -Path $dailyPath  -NoTypeInformation -Encoding UTF8
$tradesSorted | Export-Csv -Path $tradesPath -NoTypeInformation -Encoding UTF8
$added = $dailySorted.Count - $before
Write-Host ("台帳: {0}行 (今回追加 {1}行) / 建玉 {2}行" -f $dailySorted.Count, $added, $tradesSorted.Count)

# --- 読み物を作り直す(台帳が唯一の出どころ) ---
$yen = "+#,##0;-#,##0;0"
$days = @($dailySorted | Select-Object -ExpandProperty buy_date -Unique | Sort-Object)
$md = New-Object System.Collections.Generic.List[string]
$md.Add("# 未知データでの検証結果")
$md.Add("")
if ($days.Count -eq 0) {
    $md.Add("まだ記録がない。")
} else {
    $md.Add("モデルを固めた後に出てきたデータ($StartDate に買った分から)での成績。")
    $md.Add("**$($days[0]) 〜 $($days[-1]) / $($days.Count)営業日**。元手 $('{0:N0}' -f $Account)円・利益を再投資・コスト$CostBps bp・税引後。")
    $md.Add("")
    $md.Add("引けで買って翌営業日の寄りで売る。数字は一度書いたら書き換えない(``reports/forward_test/daily.csv`` が元)。")
    $md.Add("")
    $md.Add("## まとめ")
    $md.Add("")
    $md.Add("| 買い方 | 日数 | 勝ち | 勝率 | 累計損益 | 平均/日 | 最大の勝ち | 最大の負け |")
    $md.Add("|---|---:|---:|---:|---:|---:|---:|---:|")
    foreach ($v in $variants) {
        $rows = @($dailySorted | Where-Object { $_.variant -eq $v.key })
        if ($rows.Count -eq 0) { continue }
        $pnls = @($rows | ForEach-Object { [int]$_.pnl })
        $wins = @($pnls | Where-Object { $_ -gt 0 }).Count
        $sum  = ($pnls | Measure-Object -Sum).Sum
        $max  = ($pnls | Measure-Object -Maximum).Maximum
        $min  = ($pnls | Measure-Object -Minimum).Minimum
        $avg  = [Math]::Round($sum / $rows.Count)
        $cells = @(
            $v.label
            $rows.Count
            $wins
            ("{0:P1}" -f ($wins / $rows.Count))
            ("{0}円" -f $sum.ToString($yen))
            ("{0}円" -f $avg.ToString($yen))
            ("{0}円" -f $max.ToString($yen))
            ("{0}円" -f $min.ToString($yen))
        )
        $md.Add("| " + ($cells -join " | ") + " |")
    }
    $md.Add("")
    $md.Add("## 日ごと")
    $md.Add("")
    $hdr = "| 買い日 | 売り日 |"
    $sep = "|---|---|"
    foreach ($v in $variants) { $hdr += " $($v.label) |"; $sep += "---:|" }
    $md.Add($hdr)
    $md.Add($sep)
    foreach ($d in $days) {
        $sell = @($dailySorted | Where-Object { $_.buy_date -eq $d })[0].sell_date
        $line = "| $d | $sell |"
        foreach ($v in $variants) {
            $r = @($dailySorted | Where-Object { $_.buy_date -eq $d -and $_.variant -eq $v.key })
            if ($r.Count -gt 0) { $line += " " + ([int]$r[0].pnl).ToString($yen) + "円 |" }
            else { $line += " — |" }
        }
        $md.Add($line)
    }
    $md.Add("")
    $md.Add("## 注意")
    $md.Add("")
    $md.Add("- 実際に売買した記録ではなく、過去データでのシミュレーション。投資判断の助言ではない。")
    $md.Add("- 銘柄の選定に使うのは買い日の**前**営業日の引けまでのデータ。買い日の引けで買い、翌営業日の寄りで売る。")
    $md.Add("- 日数が少ないうちは、勝率も損益もたまたまの範囲を出ない。")
    $md.Add("- 現在プライムに上場している銘柄だけを対象にしているので、生存者バイアスがある。")
    $md.Add("- 建玉の内訳は ``reports/forward_test/trades.csv``。")
}
$mdPath = Resolve-ProjectPath $SummaryMd
[IO.File]::WriteAllText($mdPath, (($md -join "`r`n") + "`r`n"), (New-Object Text.UTF8Encoding $false))
Write-Host "saved $SummaryMd"
