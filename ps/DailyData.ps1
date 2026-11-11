# Shared daily-update boundaries. No network or file changes when dot-sourced.
function Get-JapanNow {
    [DateTimeOffset]::UtcNow.ToOffset([TimeSpan]::FromHours(9)).DateTime
}

function Get-CompletedTradingDate {
    param([datetime]$NowJapan = (Get-JapanNow))
    if ((Test-TradingDay -Date $NowJapan.Date) -and $NowJapan.Hour -ge 17) { return $NowJapan.Date }
    Get-PrevTradingDay -Date $NowJapan.Date
}

function Resolve-DailyTarget {
    param([string]$AsOf = '', [datetime]$NowJapan = (Get-JapanNow))
    $latest = Get-CompletedTradingDate -NowJapan $NowJapan
    if (-not $AsOf) { return $latest }
    $target = [datetime]::ParseExact($AsOf, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
    if (-not (Test-TradingDay -Date $target)) { throw "Target $AsOf is not a JP trading day" }
    if ($target -gt $latest) { throw "Target $AsOf is not closed (latest completed JP session: $($latest.ToString('yyyy-MM-dd')))" }
    return $target
}

function Get-DailyLastDate {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return '' }
    $last = Get-Content -LiteralPath $Path -Tail 1
    if (-not $last) { return '' }
    return ($last -split ',')[0].Trim('"')
}

function Get-DailyLastRow {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $header = Get-Content -LiteralPath $Path -TotalCount 1
    $last = Get-Content -LiteralPath $Path -Tail 1
    if (-not $last -or $last -eq $header) { return $null }
    try { return (@($header, $last) | ConvertFrom-Csv) } catch { return $null }
}

function Get-DailyRetryDelay {
    param([int]$Attempt, [int]$BaseSeconds)
    return $BaseSeconds * $Attempt
}

function Test-DailyOhlc {
    param($Row)
    foreach ($key in @('open', 'high', 'low', 'close')) {
        if ($null -eq $Row.$key) { return $false }
        try { $value = [double]$Row.$key } catch { return $false }
        if (-not ($value -gt 0) -or [double]::IsInfinity($value) -or [double]::IsNaN($value)) { return $false }
    }
    return $true
}

function Test-DailyStockRow {
    param($Row)
    if (-not (Test-DailyOhlc $Row)) { return $false }
    try { $volume = [double]$Row.volume } catch { return $false }
    return $volume -gt 0 -and -not [double]::IsInfinity($volume)
}

function Save-DailyDiagnostics {
    param([string]$Path, [object[]]$Attempts)
    $parent = Split-Path -Parent $Path
    New-Item -ItemType Directory -Force -Path $parent | Out-Null
    $json = ConvertTo-Json -InputObject @($Attempts) -Depth 6
    [IO.File]::WriteAllText($Path, $json, (New-Object Text.UTF8Encoding $false))
}

function Get-MarketDateLimit {
    param([string]$AsOf, [int]$Offset)
    # A JP as-of is the JP close. US/FX bars bearing that same date may still be open.
    # Keep the conservative preceding calendar date for non-JP series.
    if ($Offset -eq 9) { return $AsOf }
    return ([datetime]$AsOf).AddDays(-1).ToString('yyyy-MM-dd')
}

function Test-MarketLastDate {
    param([string]$Latest, [string]$AsOf, [int]$Offset)
    if (-not $Latest) { return $false }
    $limit = Get-MarketDateLimit -AsOf $AsOf -Offset $Offset
    if ($Offset -eq 9) { return $Latest -eq $limit }
    # Non-JP calendars differ from JP. Bound their age explicitly rather than
    # accepting arbitrarily old history; this is not an exact US/FX session check.
    $oldest = ([datetime]$limit).AddDays(-7).ToString('yyyy-MM-dd')
    return $Latest -le $limit -and $Latest -ge $oldest
}

function ConvertFrom-MarketChart {
    param($Response, [int]$Offset, [string]$LastDate)
    if ($Response.chart.error) { throw "Yahoo chart error: $($Response.chart.error.code): $($Response.chart.error.description)" }
    if (-not $Response.chart.result) { throw 'Yahoo chart result is empty' }
    $r = $Response.chart.result[0]
    if (-not $r.timestamp -or -not $r.indicators.quote) { throw 'Yahoo timestamps/quotes are empty' }
    $q = $r.indicators.quote[0]
    $adj = $null
    if ($r.indicators.adjclose) { $adj = $r.indicators.adjclose[0].adjclose }
    for ($i = 0; $i -lt $r.timestamp.Count; $i++) {
        $utc = [DateTimeOffset]::FromUnixTimeSeconds([int64]$r.timestamp[$i]).UtcDateTime
        $date = $utc.AddHours($Offset).ToString('yyyy-MM-dd')
        if ($LastDate -and $date -gt $LastDate) { continue }
        $raw = [pscustomobject]@{ open=$q.open[$i]; high=$q.high[$i]; low=$q.low[$i]; close=$q.close[$i] }
        if (-not (Test-DailyOhlc $raw)) { continue }
        $ratio = 1.0
        if ($adj -and $null -ne $adj[$i]) { $ratio = [double]$adj[$i] / [double]$raw.close }
        $row = [pscustomobject]@{
            date=$date; utc_time=$utc.ToString('yyyy-MM-dd HH:mm')
            open=[double]$raw.open*$ratio; high=[double]$raw.high*$ratio
            low=[double]$raw.low*$ratio; close=[double]$raw.close*$ratio
            raw_open=[double]$raw.open; raw_close=[double]$raw.close
            volume=$(if ($null -ne $q.volume[$i]) { [double]$q.volume[$i] } else { 0.0 })
        }
        if (Test-DailyOhlc $row) { $row }
    }
}
