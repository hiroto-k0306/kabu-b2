# 銘柄数の多いユニバース(東証プライム等)向けの株価取得。
# 銘柄ごとに data.tickerDir/<code>.csv へ保存するので、途中で止まっても再実行すれば続きから取得する。
# 1ファイルに調整済み価格(open/high/low/close/volume)と、分割・配当調整前の実際の価格(raw_*)の両方を持つ。
# 使い方: $env:KABU_CONFIG = "ps/config.prime.json"; powershell -File ps\Fetch-UniversePrices.ps1

param(
    [switch]$Refresh,
    [string]$AsOf = "",
    [ValidateRange(1, 5)][int]$MaxAttempts = 3,
    [ValidateRange(0, 600)][int]$RetryDelaySeconds = 60,
    [string]$LogDir = "logs"
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"
. "$PSScriptRoot\DailyData.ps1"
if ($AsOf) { $AsOf = (Resolve-DailyTarget -AsOf $AsOf).ToString('yyyy-MM-dd') }
$config = Get-Config
$universe = @(Import-Csv -Path (Resolve-ProjectPath $config.universe.csv) -Encoding UTF8)
$tickerDir = Split-Path (Resolve-ProjectPath "$($config.data.tickerDir)/placeholder") -Parent
$diagnostics = New-Object System.Collections.Generic.List[object]
$diagnosticPath = Resolve-ProjectPath "$LogDir/stock_fetch.json"
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$fetched = 0; $skipped = 0
$pending = @($universe | Where-Object {
    $path = Join-Path $tickerDir "$($_.code).csv"
    $cached = Get-DailyLastRow $path
    if (-not $Refresh -and (Test-Path $path) -and (Get-Item $path).Length -gt 0 -and
        (-not $AsOf -or ($cached -and $cached.date -eq $AsOf -and (Test-DailyStockRow $cached)))) { $skipped++; $false } else { $true }
})
for ($attempt = 1; $attempt -le $MaxAttempts -and $pending.Count -gt 0; $attempt++) {
    if ($attempt -gt 1) {
        $delay = Get-DailyRetryDelay -Attempt ($attempt - 1) -BaseSeconds $RetryDelaySeconds
        Write-Host "retry $attempt/$MaxAttempts: $($pending.Count) incomplete symbols after $delay seconds"
        Start-Sleep -Seconds $delay
    }
    $next = New-Object System.Collections.Generic.List[object]
    foreach ($u in $pending) {
        $path = Join-Path $tickerDir "$($u.code).csv"
        $startDate = if ($config.data.startDate) { [string]$config.data.startDate } else { '' }
        $rows = @(); $fetchError = ''
        try {
            $rows = @(Get-YahooChart -Code $u.code -Years $config.data.historyYears -WithRawPrices -StartDate $startDate -EndDate $AsOf)
        } catch { $fetchError = $_.Exception.Message }
        $last = if ($rows.Count) { [string]$rows[-1].date } else { '' }
        $validTarget = @($rows | Where-Object { $_.date -eq $AsOf -and (Test-DailyStockRow $_) }).Count -gt 0
        $ok = $rows.Count -gt 0 -and (-not $AsOf -or ($last -eq $AsOf -and $validTarget))
        $reason = if ($fetchError) { "payload_error: $fetchError" } elseif ($ok) { 'complete' } elseif (-not $rows.Count) { 'empty_or_http_failure' } else { 'missing_valid_target_bar' }
        $diagnostics.Add([pscustomobject]@{ symbol="$($u.code).T"; target=$AsOf; latestDate=$last; attempt=$attempt; rows=$rows.Count; ok=$ok; reason=$reason; checkedAt=[DateTimeOffset]::UtcNow.ToString('o') })
        if (-not $ok) {
            $next.Add($u)
            Write-Warning "$($u.code).T target=$AsOf latest=$last attempt=$attempt/$MaxAttempts reason=$reason; existing CSV preserved"
            continue
        }
        $tmp = "$path.tmp"
        $rows | Export-CsvNoBom -Path $tmp
        Move-Item -Path $tmp -Destination $path -Force
        $fetched++
        if ($fetched % 50 -eq 0) { Write-Host "fetched $fetched, skipped $skipped, elapsed $($sw.Elapsed)" }
        Start-Sleep -Milliseconds 200
    }
    $pending = @($next.ToArray())
    Save-DailyDiagnostics -Path $diagnosticPath -Attempts $diagnostics.ToArray()
}
Write-Host "done: fetched $fetched, already had $skipped, incomplete $($pending.Count), elapsed $($sw.Elapsed)"
# Small numbers of suspended/delisted names are still handled by the existing aggregate threshold.
# Unrecovered symbols retain old files; Test-DailyUpdate -PricesOnly blocks material incompleteness.
@($pending | ForEach-Object { $_.code }) | Set-Content -Path (Join-Path $tickerDir '_failed.txt') -Encoding UTF8
