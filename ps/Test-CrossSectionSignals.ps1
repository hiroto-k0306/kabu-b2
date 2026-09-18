# README 15. の事前登録どおり、東証プライム全銘柄で「出来高急増」「夜間の勢い」の信号を検証する。
#   重い計算は ps/CrossSectionStudy.cs（Add-Type でコンパイル。Windows標準の.NETのみ）
#   -Phase Develop   : 2002〜2011年で7信号の α を出し、t が最大の1つを選ぶ（α>0 かつ t≥2 でなければ打ち切り）
#   -Phase Verify    : -Signal の信号を 2012〜2021年に適用して判定
#   -Phase Reference : -Signal の信号を 2022年以降に適用（参考）
# 指定した期間の結果だけを計算・表示する（他の期間の成績は出さない）。コスト前。
# 使い方: powershell -File ps\Test-CrossSectionSignals.ps1 -Phase Develop

param(
    [ValidateSet("Develop", "Verify", "Reference")][string]$Phase = "Develop",
    [string]$Signal = "",
    [string]$TickerDir = "data/raw/stocks/prime_since2000",
    [double]$MinTurnover = 1e8,
    [int]$MinEligible = 200,
    [int]$TopK = 10,
    # 点検用: t日の実際の終値がこの値未満の銘柄を対象から外す（0で無効）
    [double]$MinRawPrice = 0,
    # 点検用: -Signal に加えて表示する信号（例: "V1"）
    [string]$AlsoShow = ""
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"
Set-Location (Get-ProjectRoot)

$ranges = @{ Develop = @("2002-01-01", "2011-12-31"); Verify = @("2012-01-01", "2021-12-31"); Reference = @("2022-01-01", "2099-12-31") }
$periodStart = $ranges[$Phase][0]; $periodEnd = $ranges[$Phase][1]
if ($Phase -ne "Develop" -and -not $Signal) { throw "-Signal を指定してください" }

Add-Type -Path (Join-Path $PSScriptRoot "CrossSectionStudy.cs")
$outDir = Split-Path (Resolve-ProjectPath "reports/cross_section/placeholder") -Parent
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$all = [CrossSectionStudy]::Run((Resolve-Path $TickerDir).Path, (Resolve-Path "data/raw/market/daily_since2000/N225.csv").Path, $MinTurnover, $MinEligible, $TopK, (Join-Path $outDir "invalid_rows.log"), $MinRawPrice)
Write-Host "computed in $($sw.Elapsed)"

# 売りの日（t+2）までが期間内に収まる日だけを使う
$rows = @($all | Where-Object { [string]::Compare($_.Date, $periodStart) -ge 0 -and [string]::Compare($_.SellDate, $periodEnd) -le 0 })
$names = [CrossSectionStudy]::SignalNames
$useNames = $names
if ($Phase -ne "Develop" -or $Signal) { $useNames = @($Signal) }
if ($AlsoShow) { $useNames = @($useNames) + @($AlsoShow -split ",") }

$flat = foreach ($r in $rows) {
    $o = [ordered]@{ date = $r.Date; eligible = $r.Eligible; market_ew = $r.MarketEw; n225 = $r.N225 }
    foreach ($n in $useNames) { $k = [Array]::IndexOf($names, $n); $o[$n] = $r.Signal[$k]; $o["picks_$n"] = $r.Picks[$k] }
    [PSCustomObject]$o
}
$flat | Export-Csv -Path (Join-Path $outDir "daily_${Phase}_minprice$MinRawPrice.csv") -NoTypeInformation -Encoding UTF8

function Get-Regression {
    param([object[]]$Rows, [string]$Y, [string]$X)
    $use = @($Rows | Where-Object { -not [double]::IsNaN($_.$Y) -and -not [double]::IsNaN($_.$X) })
    $n = $use.Count; $mx = 0.0; $my = 0.0
    foreach ($r in $use) { $mx += $r.$X; $my += $r.$Y }
    $mx /= $n; $my /= $n
    $sxx = 0.0; $sxy = 0.0
    foreach ($r in $use) { $dx = $r.$X - $mx; $sxx += $dx * $dx; $sxy += $dx * ($r.$Y - $my) }
    $beta = $sxy / $sxx; $alpha = $my - $beta * $mx
    $sse = 0.0
    foreach ($r in $use) { $e = $r.$Y - $alpha - $beta * $r.$X; $sse += $e * $e }
    $se = [Math]::Sqrt($sse / ($n - 2) * (1.0 / $n + $mx * $mx / $sxx))
    [PSCustomObject]@{ n = $n; alpha = $alpha; beta = $beta; alphaT = $alpha / $se }
}

$labels = @{ A1 = "出来高急増(1日)"; A2 = "出来高急増(5日)"; A3 = "急増×上昇"; A4 = "急増×下落"; B1 = "夜間の勢い(250日)"; B2 = "夜間の勢い(60日)"; B3 = "夜間−日中(250日)"; V1 = "値動き大(250日)"; B3L = "B3×値動き小" }
$mk = Get-SeriesStats -Rows $flat -Field market_ew
$nk = Get-SeriesStats -Rows $flat -Field n225
Write-Host ""
Write-Host ("=== {0}: {1} - {2} ({3}日, 対象銘柄 平均{4:N0})  上位{5}を等分・引け→翌寄り・コスト前 ===" -f $Phase, $flat[0].date, $flat[-1].date, $flat.Count, ($flat | Measure-Object eligible -Average).Average, $TopK)
Write-Host ("  市場(対象銘柄の等分): 平均 {0,6:N2}bp/日 年率 {1,6:N1}% 最大DD {2,6:N1}%   日経平均: 平均 {3,6:N2}bp/日 年率 {4,6:N1}%" -f ($mk.meanDaily * 1e4), ($mk.annualized * 100), ($mk.maxDD * 100), ($nk.meanDaily * 1e4), ($nk.annualized * 100))
Write-Host ("  {0,-18} {1,5} {2,8} {3,7} {4,7} {5,5} {6,10} {7,6} {8,6} {9,10} {10,6}" -f "信号", "日数", "平均bp", "年率", "最大DD", "β", "α bp/日", "α t", "β日経", "α日経bp", "t")
$res = foreach ($n in $useNames) {
    $valid = @($flat | Where-Object { -not [double]::IsNaN($_.$n) })
    $s = Get-SeriesStats -Rows $valid -Field $n
    $rm = Get-Regression -Rows $valid -Y $n -X market_ew
    $rn = Get-Regression -Rows $valid -Y $n -X n225
    Write-Host ("  {0,-18} {1,5} {2,8:N2} {3,6:N1}% {4,6:N1}% {5,5:N2} {6,10:N2} {7,6:N2} {8,6:N2} {9,10:N2} {10,6:N2}" -f "$n $($labels[$n])", $valid.Count, ($s.meanDaily * 1e4), ($s.annualized * 100), ($s.maxDD * 100), $rm.beta, ($rm.alpha * 1e4), $rm.alphaT, $rn.beta, ($rn.alpha * 1e4), $rn.alphaT)
    [PSCustomObject]@{ name = $n; alpha = $rm.alpha; t = $rm.alphaT }
}

Write-Host ""
if ($Phase -eq "Develop" -and $Signal) {
    # 1つの信号だけを事前登録した場合（README 22. など）
    $r = $res[0]
    Write-Host ("  判定: {0}" -f $(if ($r.alpha -gt 0 -and $r.t -ge 2) { "α>0 かつ t≥2 → 検証へ" } else { "基準（α>0 かつ t≥2）を満たさない → 打ち切り" }))
} elseif ($Phase -eq "Develop") {
    $best = @($res | Sort-Object t -Descending)[0]
    if ($best.alpha -gt 0 -and $best.t -ge 2) { Write-Host "  選択: $($best.name)（α t が最大、α>0 かつ t≥2）→ 段階2へ" }
    else { Write-Host ("  t が最大の信号 {0}（α {1:N2}bp, t {2:N2}）は基準（α>0 かつ t≥2）を満たさない → 打ち切り" -f $best.name, ($best.alpha * 1e4), $best.t) }
} elseif ($Phase -eq "Verify") {
    $r = $res[0]
    Write-Host ("  判定: {0}" -f $(if ($r.alpha -gt 0 -and $r.t -ge 2) { "有効（α>0 かつ t≥2）" } else { "効果は確認できない" }))
}

Write-Host ""
Write-Host "  年別（α bp/日 と t、市場=対象銘柄の等分）"
foreach ($yg in ($flat | Group-Object { $_.date.Substring(0, 4) } | Sort-Object Name)) {
    $g = @($yg.Group)
    $m = Get-SeriesStats -Rows $g -Field market_ew
    $parts = foreach ($n in $useNames) { $rg = Get-Regression -Rows $g -Y $n -X market_ew; "{0} {1,6:N2} ({2,5:N2})" -f $n, ($rg.alpha * 1e4), $rg.alphaT }
    Write-Host ("  {0} 市場 {1,6:N1}% 対象{2,4:N0}銘柄 | {3}" -f $yg.Name, ($m.totalReturn * 100), ($g | Measure-Object eligible -Average).Average, ($parts -join " | "))
}
