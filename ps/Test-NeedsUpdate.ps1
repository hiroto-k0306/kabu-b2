# GitHub Actions の最初に呼ぶ。直近の営業日の分がもう git に入っているかを、
# 追跡しているファイル(N225.csv と today_picks.json)だけで判定する。
# 毎回まっさらな環境で動くので、銘柄CSVの更新時刻は見ない(Test-DailyUpdate.ps1 との違い)。
# 結果は GITHUB_OUTPUT に needed=true/false と target=yyyy-MM-dd で書く(ローカルでは表示だけ)。
# 使い方: pwsh -File ps/Test-NeedsUpdate.ps1 [-Force]
param(
    [switch]$Force   # 最新でも更新する(手動実行用)。取引時間中は -Force でも更新しない
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"
$root = Get-ProjectRoot

# 17時以降なら当日、それより前なら前営業日の分がそろっていてほしい(Test-DailyUpdate.ps1 と同じ基準)
$now = Get-Date
if ((Test-TradingDay -Date $now.Date) -and $now.Hour -ge 17) { $target = $now.Date }
else { $target = Get-PrevTradingDay -Date $now.Date }
$targetStr = $target.ToString("yyyy-MM-dd")

function Get-LastDate {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return "" }
    $last = Get-Content -Path $Path -Tail 1
    if (-not $last) { return "" }
    return ($last -split ",")[0].Trim('"')
}

$n225Last = Get-LastDate (Join-Path $root "data/raw/market/daily_since2000/N225.csv")
$picksAsOf = ""
$picksPath = Join-Path $root "reports/today_picks.json"
if (Test-Path $picksPath) { $picksAsOf = ([IO.File]::ReadAllText($picksPath) | ConvertFrom-Json).asOf }

$inMarket = (Test-TradingDay -Date $now.Date) -and $now.Hour -ge 9 -and $now.Hour -lt 16
$upToDate = ($n225Last -eq $targetStr) -and ($picksAsOf -eq $targetStr)
Write-Host ("now {0}  target {1}  N225.csv {2}  today_picks.asOf {3}" -f $now.ToString("yyyy-MM-dd HH:mm"), $targetStr, $n225Last, $picksAsOf)

if ($inMarket) {
    # 取引時間中に取ると最新行が途中の値になる。Actions の起動が遅れて寄り付き後にずれ込んだとき用
    $needed = $false; $reason = "取引時間中なので更新しない"
} elseif ($upToDate -and -not $Force) {
    $needed = $false; $reason = "$targetStr の分はそろっている"
} else {
    $needed = $true; $reason = $(if ($upToDate) { "Force が指定された" } else { "$targetStr の分がない" })
}
Write-Host "needed=$needed ($reason)"

if ($env:GITHUB_OUTPUT) {
    Add-Content -Path $env:GITHUB_OUTPUT -Value ("needed={0}" -f $needed.ToString().ToLower())
    Add-Content -Path $env:GITHUB_OUTPUT -Value "target=$targetStr"
}
