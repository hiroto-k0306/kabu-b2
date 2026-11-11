# GitHub Actions の最初に呼ぶ。直近の営業日の分がもう git に入っているかを、
# 追跡しているファイル(N225.csv と today_picks.json)だけで判定する。
# 毎回まっさらな環境で動くので、銘柄CSVの更新時刻は見ない(Test-DailyUpdate.ps1 との違い)。
# 結果は GITHUB_OUTPUT に needed=true/false と target=yyyy-MM-dd で書く(ローカルでは表示だけ)。
# 使い方: pwsh -File ps/Test-NeedsUpdate.ps1 [-Force]
param(
    [switch]$Force   # 最新でも確定済み営業日までを更新する(手動実行用)
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"
. "$PSScriptRoot\DailyData.ps1"
$root = Get-ProjectRoot

# 17時以降なら当日、それより前なら前営業日の分がそろっていてほしい(Test-DailyUpdate.ps1 と同じ基準)
$now = Get-JapanNow
$target = Resolve-DailyTarget -NowJapan $now
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

$upToDate = ($n225Last -eq $targetStr) -and ($picksAsOf -eq $targetStr)
Write-Host ("now {0}  target {1}  N225.csv {2}  today_picks.asOf {3}" -f $now.ToString("yyyy-MM-dd HH:mm"), $targetStr, $n225Last, $picksAsOf)

# A delayed morning run still repairs the fixed previous session. Fetch scripts
# enforce the target; skipping merely because the next session opened is unnecessary.
if ($upToDate -and -not $Force) {
    $needed = $false; $reason = "$targetStr の分はそろっている"
} else {
    $needed = $true; $reason = $(if ($upToDate) { "Force が指定された" } else { "$targetStr の分がない" })
}
Write-Host "needed=$needed ($reason)"

if ($env:GITHUB_OUTPUT) {
    Add-Content -Path $env:GITHUB_OUTPUT -Value ("needed={0}" -f $needed.ToString().ToLower())
    Add-Content -Path $env:GITHUB_OUTPUT -Value "target=$targetStr"
}

