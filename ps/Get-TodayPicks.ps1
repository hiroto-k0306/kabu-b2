# 最新のデータ（判断日 t の引けまで）で、各買い方の「次の営業日の引けで買う銘柄」を出す。
#   B2（夜間の勢い×値動きの小ささ）3/5/10銘柄、日経225の売買代金上位10、東証プライムの売買代金上位10、半々（各5）
#   株数の目安は「予算 ÷ 銘柄数 ÷ 上限価格（終値＋制限値幅）」。実際の注文は成行で、余力は上限価格で拘束される
# 出力: reports/today_picks.json と画面表示
# 使い方: powershell -File ps\Get-TodayPicks.ps1 [-Budget 500000]

param(
    [double]$Budget = 500000,
    [string]$TickerDir = "data/raw/stocks/prime_since2000",
    [double]$MinTurnover = 1e8,
    [int]$MinEligible = 200
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"
Set-Location (Get-ProjectRoot)
Add-Type -Path (Join-Path $PSScriptRoot "CrossSectionStudy.cs")

$limitThresholds = [double[]]@(100, 200, 500, 700, 1000, 1500, 2000, 3000, 5000, 7000, 10000, 15000, 20000, 30000, 50000, 70000, 100000, 150000, 200000, 300000, 500000, 700000, 1000000)
$limitWidths = [double[]]@(30, 50, 80, 100, 150, 300, 400, 500, 700, 1000, 1500, 3000, 4000, 5000, 7000, 10000, 15000, 30000, 40000, 50000, 70000, 100000, 150000)
function Get-LimitWidth {
    param([double]$Base)
    for ($k = 0; $k -lt $limitThresholds.Length; $k++) { if ($Base -lt $limitThresholds[$k]) { return $limitWidths[$k] } }
    return 300000.0
}

$names = @{}; $sectors = @{}
foreach ($u in (Import-Csv data/raw/universe/prime.csv -Encoding UTF8)) { $names[$u.code] = $u.name; $sectors[$u.code] = $u.sector }

# 東証の休業日（土日＋祝日＋年末年始12/31〜1/3）。祝日は法律の規則から作った内蔵リストで、念のため取引所のカレンダーで確認すること
$holidays = @(
    "2026-01-01", "2026-01-02", "2026-01-12", "2026-02-11", "2026-02-23", "2026-03-20",
    "2026-04-29", "2026-05-03", "2026-05-04", "2026-05-05", "2026-05-06", "2026-07-20",
    "2026-08-11", "2026-09-21", "2026-09-22", "2026-09-23", "2026-10-12", "2026-11-03",
    "2026-11-23", "2026-12-31",
    "2027-01-01", "2027-01-02", "2027-01-03", "2027-01-11", "2027-02-11", "2027-02-23",
    "2027-03-21", "2027-03-22", "2027-04-29", "2027-05-03", "2027-05-04", "2027-05-05",
    "2027-07-19", "2027-08-11", "2027-09-20", "2027-09-23", "2027-10-11", "2027-11-03",
    "2027-11-23", "2027-12-31"
)
$dowJa = @{ "Sunday" = "日"; "Monday" = "月"; "Tuesday" = "火"; "Wednesday" = "水"; "Thursday" = "木"; "Friday" = "金"; "Saturday" = "土" }
function Get-NextTradingDay {
    param([string]$Date)
    $d = [datetime]$Date
    for ($i = 1; $i -le 30; $i++) {
        $x = $d.AddDays($i)
        if ($x.DayOfWeek -eq "Saturday" -or $x.DayOfWeek -eq "Sunday") { continue }
        if ($holidays -contains $x.ToString("yyyy-MM-dd")) { continue }
        return $x
    }
    return $d.AddDays(1)
}
function Format-Day {
    param([datetime]$D)
    "{0}月{1}日({2})" -f $D.Month, $D.Day, $dowJa[$D.DayOfWeek.ToString()]
}

# --- B2 の候補（最新日までのデータだけで計算） ---
$outDir = Split-Path (Resolve-ProjectPath "reports/placeholder") -Parent
$latest = [CrossSectionStudy]::LatestPicks((Resolve-Path $TickerDir).Path, (Resolve-Path "data/raw/market/daily_since2000/N225.csv").Path, $MinTurnover, $MinEligible, 20, (Join-Path $outDir "today_invalid_rows.log"))
$asOf = $latest[0]
$b2 = foreach ($s in ($latest | Select-Object -Skip 1)) {
    $f = $s -split ":"
    [PSCustomObject]@{ code = $f[0]; name = $names[$f[0]]; sector = $sectors[$f[0]]; score = [double]$f[1]; close = [double]$f[2]; vol = [double]$f[3] }
}
Write-Host "判断日（データの最終日）: $asOf"

# --- 売買代金の上位（その日のランキングから） ---
function Get-TurnoverTop {
    param([string]$Csv, [int]$K)
    $rows = @(Import-Csv $Csv | Where-Object { $_.date -eq $asOf -and [int]$_.rank -le $K } | Sort-Object { [int]$_.rank })
    foreach ($r in $rows) {
        [PSCustomObject]@{ code = $r.code; name = $names[$r.code]; sector = $sectors[$r.code]; close = [double]$r.close; rank = [int]$r.rank }
    }
}
$nk10 = @(Get-TurnoverTop -Csv "data/processed/nikkei_pit/daily_turnover_ranking_5y.csv" -K 10)
$pr10 = @(Get-TurnoverTop -Csv "data/processed/prime/daily_turnover_ranking_5y.csv" -K 10)
$nk5 = @($nk10 | Select-Object -First 5)

function Build-List {
    param($Items, [int]$K, [double]$Share)
    $sel = @($Items | Select-Object -First $K)
    $per = $Budget * $Share / $K
    foreach ($x in $sel) {
        $lock = $x.close + (Get-LimitWidth -Base $x.close)
        $sh = [int][Math]::Floor($per / $lock)
        [PSCustomObject]@{ code = $x.code; name = $x.name; sector = $x.sector
            close = [math]::Round($x.close, 1); lock = [math]::Round($lock, 1); shares = $sh
            amount = [math]::Round($sh * $x.close); budget = [math]::Round($per) }
    }
}

$strategies = [ordered]@{
    "B2 3銘柄"      = (Build-List -Items $b2 -K 3 -Share 1.0)
    "B2 5銘柄"      = (Build-List -Items $b2 -K 5 -Share 1.0)
    "B2 10銘柄"     = (Build-List -Items $b2 -K 10 -Share 1.0)
    "日経225 上位10" = (Build-List -Items $nk10 -K 10 -Share 1.0)
    "プライム 上位10" = (Build-List -Items $pr10 -K 10 -Share 1.0)
    "日経225上位5＋B2上位5" = (@(Build-List -Items $nk5 -K 5 -Share 0.5) + @(Build-List -Items $b2 -K 5 -Share 0.5))
}

foreach ($k in $strategies.Keys) {
    $list = @($strategies[$k])
    $total = 0.0; foreach ($x in $list) { $total += $x.amount }
    Write-Host ""
    Write-Host ("=== {0}（予算 {1:N0}円 / 概算の買付額 {2:N0}円） ===" -f $k, $Budget, $total)
    Write-Host ("  {0,-6} {1,-22} {2,9} {3,9} {4,6} {5,10}" -f "コード", "銘柄", "終値", "上限価格", "株数", "概算金額")
    foreach ($x in $list) {
        Write-Host ("  {0,-6} {1,-22} {2,9:N1} {3,9:N1} {4,6} {5,10:N0}{6}" -f $x.code, $x.name, $x.close, $x.lock, $x.shares, $x.amount, $(if ($x.shares -eq 0) { "  ← 予算内で買えない" } else { "" }))
    }
}

$buyDay = Get-NextTradingDay -Date $asOf
$sellDay = Get-NextTradingDay -Date $buyDay.ToString("yyyy-MM-dd")
$holdDays = ($sellDay - $buyDay).TotalDays
Write-Host ""
Write-Host ("買い: {0} 10:30〜14:00 に成行 → 当日15:30の終値で約定" -f (Format-Day $buyDay))
Write-Host ("売り: {0} 14:00 〜 {1} 7:00 に成行 → {1} 9:00の始値で約定（保有 {2} 日）" -f (Format-Day $buyDay), (Format-Day $sellDay), $holdDays)

$json = [PSCustomObject]@{
    asOf = $asOf; budget = $Budget; generatedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm")
    buyDate = $buyDay.ToString("yyyy-MM-dd"); buyDateLabel = (Format-Day $buyDay)
    sellDate = $sellDay.ToString("yyyy-MM-dd"); sellDateLabel = (Format-Day $sellDay)
    holdDays = $holdDays
    strategies = $strategies
} | ConvertTo-Json -Depth 6 -Compress
[IO.File]::WriteAllText((Resolve-ProjectPath "reports/today_picks.json"), $json, (New-Object Text.UTF8Encoding $false))
Write-Host ""
Write-Host "saved reports/today_picks.json"
