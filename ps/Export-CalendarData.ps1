# Simulate-SKabu.ps1 の出力(trades_*.csv / daily_*.csv)を web/calendar.html 用の calendar_data.json にまとめる。
#   days[]  : 買った日ごとの建玉。pnl は手数料・コスト込み(tradesのpnlをそのまま合算)
#   equity  : 日付 -> 税引後の評価額(daily.csv の equity)
# 先に 6通りの設定で Simulate-SKabu.ps1 を実行しておく(ps\Update-B3L2026.ps1 がまとめて行う)。
# 使い方: powershell -File ps\Export-CalendarData.ps1
param(
    [string]$ReportRoot = "reports/b3l_2026",
    [string]$OutJson    = "reports/b3l_2026/calendar_data.json",
    # シミュレーションは助走のため開始前から価格を読むが、カレンダーに出すのは運用開始以降だけ
    [string]$StartDate  = "2026-01-01",
    [int]$Account       = 500000,
    [int]$CostBps       = 3
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"

Set-Location (Get-ProjectRoot)

# key, 出力ディレクトリ, シナリオ名, 表示名。順序が calendar.html のタブ順になる
$variants = @(
    @{ key = "b2_3";    dir = "top3";    scenario = "top3_compound";    label = "B2 3銘柄" }
    @{ key = "b2_5";    dir = "top5";    scenario = "top5_compound";    label = "B2 5銘柄" }
    @{ key = "b2_10";   dir = "top10";   scenario = "top10_compound";   label = "B2 10銘柄" }
    @{ key = "nk10";    dir = "nk10";    scenario = "nk10_compound";    label = "日経225 上位10" }
    @{ key = "prime10"; dir = "prime10"; scenario = "prime10_compound"; label = "プライム 上位10" }
    @{ key = "mix55";   dir = "mix55";   scenario = "mix55_compound";   label = "日経225上位5＋B2上位5" }
)

# 銘柄名
$names = @{}
foreach ($u in (Import-Csv -Path (Resolve-ProjectPath "data/raw/universe/prime.csv") -Encoding UTF8)) { $names[[string]$u.code] = [string]$u.name }

$variantOut = [ordered]@{}
$order = New-Object System.Collections.Generic.List[string]
$monthSet = New-Object System.Collections.Generic.HashSet[string]

foreach ($v in $variants) {
    $tag       = "$($v.scenario)_cost$CostBps"
    $tradesCsv = Resolve-ProjectPath "$ReportRoot/$($v.dir)/trades_$tag.csv"
    $dailyCsv  = Resolve-ProjectPath "$ReportRoot/$($v.dir)/daily_$tag.csv"
    if (-not (Test-Path $tradesCsv)) { Write-Warning "skip $($v.key): $tradesCsv がない"; continue }
    if (-not (Test-Path $dailyCsv))  { Write-Warning "skip $($v.key): $dailyCsv がない";  continue }

    # 買った日ごとにまとめる
    $byDay = [ordered]@{}
    foreach ($t in (Import-Csv -Path $tradesCsv -Encoding UTF8)) {
        $d = [string]$t.buy_date
        if (-not $byDay.Contains($d)) {
            $byDay[$d] = [PSCustomObject]@{ date = $d; sellDate = [string]$t.sell_date; buyAmount = 0.0; sellAmount = 0.0; pnl = 0.0; fee = 0.0; trades = (New-Object System.Collections.Generic.List[object]) }
        }
        $day    = $byDay[$d]
        $shares = [int]$t.shares
        $buy    = [double]$t.buy_price
        $sell   = [double]$t.sell_price_raw
        $bAmt   = [double]$t.buy_amount
        # 実際の受取額。配当落ち・分割の日は 株数×生の売値 と一致しないが、
        # こちらでないと 売却額−買付額−手数料 = 損益 が崩れる。
        $sAmt   = [double]$t.proceeds_div_adjusted
        $pnl    = [double]$t.pnl
        $fee    = [double]$t.fee
        $code   = [string]$t.code
        $day.trades.Add([PSCustomObject]@{
            code = $code; name = $(if ($names.ContainsKey($code)) { $names[$code] } else { $code })
            # 値段は0.5円刻みがあるので整数には丸めない。
            # 取得元が単精度なため 18499.9995... のような誤差が乗るので小数1桁までにする。
            # 整数になる値は整数型にする(PowerShell 7 の ConvertTo-Json は 3285.0 と書くため)
            shares = $shares; buy = (ConvertTo-JsonNumber ([Math]::Round($buy, 1))); sell = (ConvertTo-JsonNumber ([Math]::Round($sell, 1)))
            buyAmount = [long][Math]::Round($bAmt); sellAmount = [long][Math]::Round($sAmt); pnl = [long][Math]::Round($pnl)
        })
        $day.buyAmount  += $bAmt
        $day.sellAmount += $sAmt
        $day.pnl        += $pnl
        $day.fee        += $fee
    }

    $days = New-Object System.Collections.Generic.List[object]
    foreach ($d in ($byDay.Keys | Sort-Object)) {
        $day = $byDay[$d]
        # 損益の大きい順(同じ損益はコード順)。calendar.html は並び順をそのまま表に出す
        $sorted = @($day.trades | Sort-Object @{ Expression = "pnl"; Descending = $true }, @{ Expression = "code"; Descending = $false })
        $days.Add([PSCustomObject]@{
            date = $day.date; sellDate = $day.sellDate
            buyAmount = [int][Math]::Round($day.buyAmount); sellAmount = [int][Math]::Round($day.sellAmount)
            pnl = [int][Math]::Round($day.pnl); fee = [int][Math]::Round($day.fee)
            trades = $sorted
        })
        [void]$monthSet.Add($day.date.Substring(0, 7))
    }

    $equity = [ordered]@{}
    foreach ($r in (Import-Csv -Path $dailyCsv -Encoding UTF8)) {
        if ([string]::Compare([string]$r.date, $StartDate) -lt 0) { continue }
        $equity[[string]$r.date] = [int][Math]::Round([double]$r.equity)
    }

    $variantOut[$v.key] = [PSCustomObject]@{ label = $v.label; days = $days.ToArray(); equity = $equity }
    $order.Add($v.key)
    Write-Host ("{0,-8} {1,4} days  {2} .. {3}  最終評価額 {4:N0}円" -f $v.key, $days.Count, $days[0].date, $days[-1].date, $equity[@($equity.Keys)[-1]])
}

# --- 放置パターン: 最初の買い日の引けに元手を全額投入し、以後は売買しない ---
# 投資信託（S&P500・オルカン）の基準価額は日足の取得先に無いので、連動する東証ETFの価格で代用する。
#   2558 MAXIS 米国株式(S&P500)上場投信 / 2559 MAXIS 全世界株式(オール・カントリー)上場投信
# 価格は分配金込み（配当調整済みの終値）。金額指定で全額を投入する（投資信託と同様に端数は出さない）。
# days[] は「その日の引けの評価額 → 次の営業日の引けの評価額」の値動き（売買はしない）。equity は各日の引けの評価額。
# Fetch-MarketData.ps1 が data/raw/market/daily_since2000/<コード>.csv に保存する。
function Repair-EtfSeries {
    # 取得元（Yahoo）の東証ETFの日足には、分割が調整されていない・1日だけ桁のずれた行がある
    # （2558・2559 は 2026-06-05 に10分割。6/5 から価格が10分の1になり、6/8 だけ更に10分の1の異常値が入る）。
    # 1) 前日比が 1/k、翌日比が k で元の水準に戻る1日だけの異常値の行は捨てる
    # 2) 前日比が 1/k（k=2〜20、誤差15%以内）の行は分割とみなし、それ以前の価格を k で割る
    # 取得元が直したときは何もしない（跳びが無ければ補正しない）。補正したときは内容を表示する。
    param([string[]]$Dates, [double[]]$Close, [string]$Code)
    $d = New-Object System.Collections.Generic.List[string]
    $c = New-Object System.Collections.Generic.List[double]
    for ($i = 0; $i -lt $Close.Count; $i++) { $d.Add($Dates[$i]); $c.Add($Close[$i]) }
    $i = 1
    while ($i -lt $c.Count - 1) {
        $r1 = $c[$i] / $c[$i - 1]
        $back = $c[$i + 1] / $c[$i - 1]
        if (($r1 -lt 0.6 -or $r1 -gt 1.6) -and $back -gt 0.8 -and $back -lt 1.25) {
            Write-Host ("  {0}: {1} の終値 {2} は異常値（前後は {3} と {4}）なので使わない" -f $Code, $d[$i], $c[$i], $c[$i - 1], $c[$i + 1])
            $d.RemoveAt($i); $c.RemoveAt($i)
        } else { $i++ }
    }
    for ($i = 1; $i -lt $c.Count; $i++) {
        $r = $c[$i] / $c[$i - 1]
        if ($r -lt 0.55) {
            $k = [Math]::Round(1.0 / $r)
            if ($k -ge 2 -and $k -le 20 -and [Math]::Abs($r * $k - 1.0) -lt 0.15) {
                Write-Host ("  {0}: {1} に {2}分割とみなし、それ以前の価格を {2} で割る" -f $Code, $d[$i], $k)
                for ($j = 0; $j -lt $i; $j++) { $c[$j] = $c[$j] / $k }
            }
        }
    }
    return [PSCustomObject]@{ dates = $d.ToArray(); close = $c.ToArray() }
}

$holds = @(
    @{ key = "hold_sp500"; code = "2558"; name = "MAXIS 米国株式(S&P500)上場投信"; label = "S&P500 放置（2558）" }
    @{ key = "hold_acwi";  code = "2559"; name = "MAXIS 全世界株式(オール・カントリー)上場投信"; label = "オルカン放置（2559）" }
)
$firstBuy = ""
foreach ($k in $variantOut.Keys) {
    $d0 = [string]$variantOut[$k].days[0].date
    if ($firstBuy -eq "" -or [string]::Compare($d0, $firstBuy) -lt 0) { $firstBuy = $d0 }
}
foreach ($h in $holds) {
    $csv = Resolve-ProjectPath "data/raw/market/daily_since2000/$($h.code).csv"
    if (-not (Test-Path $csv)) { Write-Warning "skip $($h.key): $csv がない"; continue }
    $raw = @(Import-Csv -Path $csv -Encoding UTF8)
    $fix = Repair-EtfSeries -Dates ([string[]]@($raw | ForEach-Object { [string]$_.date })) -Close ([double[]]@($raw | ForEach-Object { [double]$_.close })) -Code $h.code
    $px = @()
    for ($i = 0; $i -lt $fix.dates.Length; $i++) {
        if ([string]::Compare($fix.dates[$i], $StartDate) -ge 0) { $px += [PSCustomObject]@{ date = $fix.dates[$i]; close = $fix.close[$i] } }
    }
    $i0 = -1
    for ($i = 0; $i -lt $px.Count; $i++) { if ([string]$px[$i].date -eq $firstBuy) { $i0 = $i; break } }
    if ($i0 -lt 0) { Write-Warning "skip $($h.key): $firstBuy の価格がない"; continue }
    $units = $Account / [double]$px[$i0].close
    $hDays = New-Object System.Collections.Generic.List[object]
    $hEquity = [ordered]@{}
    for ($i = 0; $i -lt $i0; $i++) { $hEquity[[string]$px[$i].date] = [int]$Account }
    $hEquity[$firstBuy] = [int]$Account
    for ($i = $i0; $i -lt $px.Count - 1; $i++) {
        $d = [string]$px[$i].date; $n = [string]$px[$i + 1].date
        $c0 = [double]$px[$i].close; $c1 = [double]$px[$i + 1].close
        $bAmt = $units * $c0; $sAmt = $units * $c1; $pnl = $sAmt - $bAmt
        $hDays.Add([PSCustomObject]@{
            date = $d; sellDate = $n
            buyAmount = [int][Math]::Round($bAmt); sellAmount = [int][Math]::Round($sAmt)
            pnl = [int][Math]::Round($pnl); fee = 0
            trades = @([PSCustomObject]@{
                code = $h.code; name = $h.name
                shares = [int][Math]::Round($units); buy = (ConvertTo-JsonNumber ([Math]::Round($c0, 1))); sell = (ConvertTo-JsonNumber ([Math]::Round($c1, 1)))
                buyAmount = [long][Math]::Round($bAmt); sellAmount = [long][Math]::Round($sAmt); pnl = [long][Math]::Round($pnl)
            })
        })
        $hEquity[$n] = [int][Math]::Round($sAmt)
        [void]$monthSet.Add($d.Substring(0, 7))
    }
    if ($hDays.Count -eq 0) { Write-Warning "skip $($h.key): 日数が足りない"; continue }
    $variantOut[$h.key] = [PSCustomObject]@{ label = $h.label; kind = "hold"; days = $hDays.ToArray(); equity = $hEquity }
    $order.Add($h.key)
    Write-Host ("{0,-10} {1,4} days  {2} .. {3}  最終評価額 {4:N0}円" -f $h.key, $hDays.Count, $hDays[0].date, $hDays[-1].date, $hEquity[@($hEquity.Keys)[-1]])
}

if ($order.Count -eq 0) { throw "シミュレーション結果が1つも見つからない" }

$out = [PSCustomObject]@{
    account  = $Account
    costBps  = $CostBps
    compound = $true
    months   = @($monthSet | Sort-Object)
    order    = $order.ToArray()
    variants = $variantOut
}

$json = $out | ConvertTo-Json -Depth 8 -Compress
$path = Resolve-ProjectPath $OutJson
New-Item -ItemType Directory -Force (Split-Path $path -Parent) | Out-Null
[IO.File]::WriteAllText($path, $json, (New-Object Text.UTF8Encoding $false))
Write-Host "saved $OutJson ($([Math]::Round($json.Length / 1KB))KB)"
