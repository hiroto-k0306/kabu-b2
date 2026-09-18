# タイミング信号の検証用に、指数・ETF・為替の日足を取得できる限り過去までさかのぼって保存する。
# 保存先: data/raw/market/daily_since2000/<名前>.csv（date は取引所の現地日付、open/high/low/close は配当調整済み）
# 使い方: powershell -File ps\Fetch-MarketData.ps1

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"
Set-Location (Get-ProjectRoot)

# 名前, Yahooのシンボル, 足の時刻に足す時間（現地日付にするため。為替の足はロンドン0時=UTC23時/0時に始まるので+1）
$targets = @(
    @{ name = "1321"; symbol = "1321.T"; offset = 9 },
    @{ name = "N225"; symbol = "^N225"; offset = 9 },
    @{ name = "1571"; symbol = "1571.T"; offset = 9 },   # 日経平均インバース（-1倍）
    @{ name = "1357"; symbol = "1357.T"; offset = 9 },   # 日経ダブルインバース（-2倍）
    @{ name = "GSPC"; symbol = "^GSPC"; offset = -5 },
    @{ name = "VIX"; symbol = "^VIX"; offset = -5 },
    @{ name = "USDJPY"; symbol = "JPY=X"; offset = 1 }
)
$outDir = Split-Path (Resolve-ProjectPath "data/raw/market/daily_since2000/placeholder") -Parent
$headers = @{ "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" }

foreach ($tg in $targets) {
    $uri = "https://query1.finance.yahoo.com/v8/finance/chart/$([uri]::EscapeDataString($tg.symbol))?period1=946684800&period2=$([DateTimeOffset]::UtcNow.ToUnixTimeSeconds())&interval=1d"
    $resp = $null
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try { $resp = Invoke-RestMethod -Uri $uri -Headers $headers -TimeoutSec 60; break } catch { Start-Sleep -Seconds (2 * $attempt) }
    }
    if ($null -eq $resp) { Write-Warning "fetch failed: $($tg.symbol)"; continue }
    $r = $resp.chart.result[0]
    $q = $r.indicators.quote[0]
    $adj = $null
    if ($r.indicators.adjclose) { $adj = $r.indicators.adjclose[0].adjclose }
    $rows = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $r.timestamp.Count; $i++) {
        $o = $q.open[$i]; $h = $q.high[$i]; $l = $q.low[$i]; $c = $q.close[$i]
        if ($null -eq $o -or $null -eq $c -or $null -eq $h -or $null -eq $l -or [double]$c -le 0 -or [double]$o -le 0) { continue }
        $utc = [DateTimeOffset]::FromUnixTimeSeconds([int64]$r.timestamp[$i]).UtcDateTime
        $ratio = 1.0
        if ($adj -and $null -ne $adj[$i]) { $ratio = [double]$adj[$i] / [double]$c }
        $v = 0.0
        if ($null -ne $q.volume[$i]) { $v = [double]$q.volume[$i] }
        $rows.Add([PSCustomObject]@{
                date = $utc.AddHours($tg.offset).ToString("yyyy-MM-dd")
                utc_time = $utc.ToString("yyyy-MM-dd HH:mm")
                open = [double]$o * $ratio; high = [double]$h * $ratio; low = [double]$l * $ratio; close = [double]$c * $ratio
                raw_open = [double]$o; raw_close = [double]$c; volume = $v
            })
    }
    $path = Join-Path $outDir "$($tg.name).csv"
    $rows | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
    Write-Host ("{0,-7} {1,5} rows  {2} - {3}  (UTC time of first/last bar: {4} / {5})" -f $tg.name, $rows.Count, $rows[0].date, $rows[-1].date, $rows[0].utc_time, $rows[-1].utc_time)
    Start-Sleep -Milliseconds 500
}
