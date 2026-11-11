# タイミング信号の検証用に、指数・ETF・為替の日足を取得できる限り過去までさかのぼって保存する。
# 保存先: data/raw/market/daily_since2000/<名前>.csv（date は取引所の現地日付、open/high/low/close は配当調整済み）
# 使い方: powershell -File ps\Fetch-MarketData.ps1

param(
    [string]$AsOf = "",
    [ValidateRange(1, 5)][int]$MaxAttempts = 3,
    [ValidateRange(0, 600)][int]$RetryDelaySeconds = 60,
    [string]$LogDir = "logs"
)
$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"
. "$PSScriptRoot\DailyData.ps1"
$AsOf = (Resolve-DailyTarget -AsOf $AsOf).ToString('yyyy-MM-dd')
Set-Location (Get-ProjectRoot)

# 名前, Yahooのシンボル, 足の時刻に足す時間（現地日付にするため。為替の足はロンドン0時=UTC23時/0時に始まるので+1）
$targets = @(
    @{ name = "1321"; symbol = "1321.T"; offset = 9 },
    @{ name = "N225"; symbol = "^N225"; offset = 9 },
    @{ name = "1571"; symbol = "1571.T"; offset = 9 },   # 日経平均インバース（-1倍）
    @{ name = "1357"; symbol = "1357.T"; offset = 9 },   # 日経ダブルインバース（-2倍）
    @{ name = "GSPC"; symbol = "^GSPC"; offset = -5 },
    @{ name = "VIX"; symbol = "^VIX"; offset = -5 },
    @{ name = "USDJPY"; symbol = "JPY=X"; offset = 1 },
    # カレンダーの「放置」パターン用（東証ETF。投資信託の基準価額は取れないので、連動するETFで代用する）
    @{ name = "2558"; symbol = "2558.T"; offset = 9 },   # MAXIS 米国株式(S&P500)上場投信
    @{ name = "2559"; symbol = "2559.T"; offset = 9 }    # MAXIS 全世界株式(オール・カントリー)上場投信
)
$outDir = Split-Path (Resolve-ProjectPath "data/raw/market/daily_since2000/placeholder") -Parent
$headers = @{ "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" }

$diagnostics = New-Object System.Collections.Generic.List[object]
$failed = New-Object System.Collections.Generic.List[string]
$diagnosticPath = Resolve-ProjectPath "$LogDir/market_fetch.json"
foreach ($tg in $targets) {
    $limit = Get-MarketDateLimit -AsOf $AsOf -Offset $tg.offset
    # Request through the inclusive cutoff; still filter returned bars defensively below.
    $p2 = ([DateTimeOffset][datetime]::SpecifyKind(([datetime]$limit).AddDays(1).AddHours(-$tg.offset), 'Utc')).ToUnixTimeSeconds()
    $uri = "https://query1.finance.yahoo.com/v8/finance/chart/$([uri]::EscapeDataString($tg.symbol))?period1=946684800&period2=$p2&interval=1d"
    $ok = $false
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        if ($attempt -gt 1) { Start-Sleep -Seconds (Get-DailyRetryDelay -Attempt ($attempt - 1) -BaseSeconds $RetryDelaySeconds) }
        $rows = @(); $latest = ''; $returnedLast = ''; $reason = ''
        try {
            $resp = Invoke-RestMethod -Uri $uri -Headers $headers -TimeoutSec 60
            if ($resp.chart.result -and $resp.chart.result[0].timestamp) {
                $stamp = $resp.chart.result[0].timestamp[-1]
                $returnedLast = [DateTimeOffset]::FromUnixTimeSeconds([int64]$stamp).UtcDateTime.AddHours($tg.offset).ToString('yyyy-MM-dd')
            }
            $rows = @(ConvertFrom-MarketChart -Response $resp -Offset $tg.offset -LastDate $limit)
            if ($rows.Count) { $latest = [string]$rows[-1].date }
            # JP benchmarks must contain the exact required session, not just an HTTP 200.
            $ok = $rows.Count -gt 0 -and (Test-MarketLastDate -Latest $latest -AsOf $AsOf -Offset $tg.offset)
            $reason = if ($ok) { 'complete' } else { 'missing_valid_target_bar' }
        } catch { $reason = "request_or_payload_error: $($_.Exception.Message)" }
        $diagnostics.Add([pscustomobject]@{ symbol=$tg.symbol; target=$AsOf; cutoff=$limit; latestReturnedDate=$returnedLast; latestValidDate=$latest; attempt=$attempt; rows=$rows.Count; ok=$ok; reason=$reason; checkedAt=[DateTimeOffset]::UtcNow.ToString('o') })
        Save-DailyDiagnostics -Path $diagnosticPath -Attempts $diagnostics.ToArray()
        if ($ok) {
            $path = Join-Path $outDir "$($tg.name).csv"
            $rows | Export-CsvNoBom -Path "$path.tmp"
            Move-Item -Path "$path.tmp" -Destination $path -Force
            Write-Host "$($tg.symbol) target=$AsOf cutoff=$limit latest=$latest rows=$($rows.Count) attempt=$attempt"
            break
        }
        Write-Warning "$($tg.symbol) target=$AsOf cutoff=$limit returned=$returnedLast latestValid=$latest attempt=$attempt/$MaxAttempts reason=$reason; existing CSV preserved"
    }
    if (-not $ok) { $failed.Add($tg.symbol) }
    Start-Sleep -Milliseconds 500
}
if ($failed.Count) { throw "Market data incomplete after $MaxAttempts attempts: $($failed -join ', '). See $diagnosticPath" }
