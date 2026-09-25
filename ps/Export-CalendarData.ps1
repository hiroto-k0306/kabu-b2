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
