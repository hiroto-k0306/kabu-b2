# 営業日の大引け後に株価を取り直し、その日のデータでカレンダーと翌営業日の銘柄まで更新する。
# タスクスケジューラから17時に呼ばれる想定。休業日は何もせずに終わる。
# 使い方: powershell -File ps\Invoke-DailyUpdate.ps1
param(
    [switch]$Force,        # 休業日でも実行する
    [switch]$PricesOnly,   # 株価の取得だけ行い、集計はしない
    [int]$Budget = 500000,
    [string]$LogDir = "logs"
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"
Set-Location (Get-ProjectRoot)
$root = Get-ProjectRoot

$today    = (Get-Date).Date
$todayStr = $today.ToString("yyyy-MM-dd")
$logDirFull = Join-Path $root $LogDir
New-Item -ItemType Directory -Force $logDirFull | Out-Null
$logPath    = Join-Path $logDirFull ("daily_{0}.log" -f $today.ToString("yyyyMMdd"))
$statusPath = Join-Path $logDirFull "last_run.json"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "{0} [{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Write-Host $line
    Add-Content -Path $logPath -Value $line -Encoding UTF8
}

$steps  = New-Object System.Collections.Generic.List[object]
$status = [ordered]@{
    date = $todayStr; startedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    finishedAt = ""; ok = $false; skipped = $false; error = ""; steps = $steps
}
function Save-Status {
    $status.finishedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $json = [PSCustomObject]$status | ConvertTo-Json -Depth 5
    [IO.File]::WriteAllText($statusPath, $json, (New-Object Text.UTF8Encoding $false))
}

Write-Log "=== 日次更新 開始 ($todayStr $(Format-JpDate -Date $today)) ==="

if (-not $Force -and -not (Test-TradingDay -Date $today)) {
    # last_run.json は上書きしない。22時/翌朝の確認が「直近に実際に動いた回」を
    # 見られなくなるため。休業日に動いたことはこのログに残る。
    Write-Log "東証の休業日なので何もしない"
    exit 0
}

# 取引時間中(9:00〜15:30)に取ると最新行が途中の値になる
$now = Get-Date
if (-not $Force -and $now.Hour -lt 16) {
    Write-Log "大引け後(16時以降)に実行すること。今は $($now.ToString('HH:mm'))" "ERROR"
    $status.error = "16時より前の実行"
    Save-Status
    exit 1
}

function Invoke-Step {
    param([string]$Title, [string]$Script, [string[]]$ScriptArgs = @(), [string]$Config = "")
    Write-Log "--- $Title"
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    if ($Config) { $env:KABU_CONFIG = $Config } else { Remove-Item Env:\KABU_CONFIG -ErrorAction SilentlyContinue }
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root "ps\$Script") @ScriptArgs *>> $logPath
    $code = $LASTEXITCODE
    $steps.Add([PSCustomObject]@{ name = $Title; script = $Script; seconds = [Math]::Round($sw.Elapsed.TotalSeconds); exitCode = $code })
    if ($code -ne 0) { throw "$Script が失敗した (exit $code)" }
    Write-Log "    完了 $($sw.Elapsed.ToString('hh\:mm\:ss'))"
}

try {
    Invoke-Step -Title "株価の取得(全銘柄を取り直す)" -Script "Fetch-UniversePrices.ps1" -ScriptArgs @("-Refresh") -Config "ps/config.prime_long.json"
    Invoke-Step -Title "指数・為替の取得"             -Script "Fetch-MarketData.ps1"

    if ($PricesOnly) {
        Write-Log "PricesOnly が指定されたので集計は行わない"
    } else {
        Invoke-Step -Title "ランキング〜カレンダー更新" -Script "Update-B3L2026.ps1" -ScriptArgs @("-Budget", "$Budget")
    }

    $status.ok = $true
    Write-Log "=== 日次更新 完了 ==="
} catch {
    $status.error = "$_"
    Write-Log "失敗: $_" "ERROR"
    Save-Status
    exit 1
}

Save-Status
exit 0
