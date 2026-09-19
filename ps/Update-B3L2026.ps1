# 株価更新後、カレンダーに出す6通りのシミュレーションをまとめて走らせ、calendar.html まで更新する。
# 前提: Fetch-UniversePrices.ps1 / Fetch-MarketData.ps1 で株価が最新になっていること。
# 使い方: powershell -File ps\Update-B3L2026.ps1
param(
    [int]$Budget = 500000,
    [switch]$SkipSignals,   # B2の選定(重い)を飛ばして、既存の picks でシミュレーションだけやり直す
    [switch]$SkipRanking    # 売買代金ランキングの作り直しを飛ばす
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"
Set-Location (Get-ProjectRoot)
$root = Get-ProjectRoot

function Invoke-Step {
    param([string]$Title, [string]$Script, [string[]]$ScriptArgs = @(), [string]$Config = "")
    Write-Host ""
    Write-Host "=== $Title ===" -ForegroundColor Cyan
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    if ($Config) { $env:KABU_CONFIG = $Config } else { Remove-Item Env:\KABU_CONFIG -ErrorAction SilentlyContinue }
    & powershell -NoProfile -File (Join-Path $root "ps\$Script") @ScriptArgs
    if ($LASTEXITCODE -ne 0) { throw "$Script が失敗した (exit $LASTEXITCODE)" }
    Write-Host ("--- $Title 完了 {0}" -f $sw.Elapsed)
}

# 1. 売買代金ランキング(日経225の各時点構成 / 東証プライム)
if (-not $SkipRanking) {
    Invoke-Step -Title "売買代金ランキング(日経225)" -Script "Build-TurnoverRanking.ps1" -Config "ps/config.nikkei_pit.json"
    Invoke-Step -Title "売買代金ランキング(プライム)" -Script "Build-TurnoverRanking.ps1" -Config "ps/config.prime.json"
}

# 2. B2の選定 -> data/processed/b3l/*.csv
if (-not $SkipSignals) {
    Invoke-Step -Title "B2の信号と上位銘柄(TopK=10)" -Script "Test-CrossSectionSignals.ps1" -ScriptArgs @("-Phase", "Reference", "-Signal", "B3L", "-TopK", "10")
    Invoke-Step -Title "picksをdate,rank,codeへ変換" -Script "Export-B3LPicks.ps1"
}

# 3. 比較用ユニバース -> data/processed/combo/*.csv
Invoke-Step -Title "比較用ユニバースの作成" -Script "Build-ComboPicks.ps1"

# 4. 6通りのS株シミュレーション
foreach ($c in @("top3", "top5", "top10", "nk10", "prime10", "mix55")) {
    Invoke-Step -Title "シミュレーション $c" -Script "Simulate-SKabu.ps1" -Config "ps/config.b3l_2026_$c.json"
}

# 5. カレンダー用データ と 翌営業日の銘柄
Invoke-Step -Title "calendar_data.json の作成" -Script "Export-CalendarData.ps1"
Invoke-Step -Title "翌営業日に買う銘柄" -Script "Get-TodayPicks.ps1" -ScriptArgs @("-Budget", "$Budget")
Invoke-Step -Title "calendar.html への埋め込み" -Script "Update-Calendar.ps1"

# 5-2. 未知データでの成績を台帳に足す(既に書いた日は触らない)
Invoke-Step -Title "未知データの検証結果を記録" -Script "Update-ForwardTest.ps1"

Write-Host ""
Write-Host "全部完了。web\calendar.html をブラウザで開く" -ForegroundColor Green
