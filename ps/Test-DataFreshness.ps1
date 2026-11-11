# Offline synthetic regression. HTTP and sleep are mocked; never fetches quotes,
# runs a strategy/backtest, touches production data, or sends an order.
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot/Common.ps1"
. "$PSScriptRoot/DailyData.ps1"
function Assert($ok, $message) { if (-not $ok) { throw $message } }
$sourceRoot = Split-Path $PSScriptRoot -Parent
$fixture = Join-Path ([IO.Path]::GetTempPath()) ('kabu-freshness-' + [guid]::NewGuid().ToString('N'))
$oldLocation = Get-Location
$oldConfig = $env:KABU_CONFIG
$oldOutput = $env:GITHUB_OUTPUT
$script:requestCounts = @{}
$script:sleeps = New-Object System.Collections.Generic.List[int]
$script:mode = 'market-recovery'
$script:target = '2026-11-10'

function New-ChartResponse([bool]$MissingTarget) {
    $stamps = @('2026-11-09', '2026-11-10', '2026-11-11') | ForEach-Object {
        ([DateTimeOffset][datetime]::SpecifyKind([datetime]$_, 'Utc')).ToUnixTimeSeconds()
    }
    $middle = if ($MissingTarget) { $null } else { 100.0 }
    [pscustomobject]@{ chart=[pscustomobject]@{ error=$null; result=@([pscustomobject]@{
        timestamp=@($stamps)
        indicators=[pscustomobject]@{ quote=@([pscustomobject]@{
            open=@(100.0,$middle,999.0); high=@(101.0,$middle,999.0)
            low=@(99.0,$middle,999.0); close=@(100.0,$middle,999.0); volume=@(1000,1000,1000)
        }); adjclose=@([pscustomobject]@{adjclose=@(100.0,$middle,999.0)}) }
    }) } }
}
function Invoke-RestMethod {
    param($Uri, $Headers, $TimeoutSec)
    # This is the only HTTP command used by the tested fetchers. No real HTTP path exists.
    $symbol = [uri]::UnescapeDataString(([uri]$Uri).AbsolutePath.Split('/')[-1])
    if (-not $script:requestCounts.ContainsKey($symbol)) { $script:requestCounts[$symbol] = 0 }
    $script:requestCounts[$symbol]++
    if ($script:mode -eq 'api-error' -and $script:requestCounts[$symbol] -eq 1) { throw 'synthetic HTTP timeout' }
    $missing = ($symbol -eq '^N225' -and ($script:mode -eq 'market-failure' -or $script:requestCounts[$symbol] -eq 1)) -or
        ($symbol -eq '1111.T' -and $script:requestCounts[$symbol] -eq 1) -or $symbol -eq '2222.T'
    if ($script:mode -eq 'malformed') { return [pscustomobject]@{chart=[pscustomobject]@{error=[pscustomobject]@{code='Bad';description='synthetic'}}} }
    New-ChartResponse $missing
}
function Start-Sleep {
    param([int]$Seconds=0, [int]$Milliseconds=0)
    if ($Seconds -gt 0) { $script:sleeps.Add($Seconds) }
}
function Write-PriceFile([string]$Path, [string]$Date) {
    New-Item -ItemType Directory -Force (Split-Path -Parent $Path) | Out-Null
    [IO.File]::WriteAllLines($Path, @('date,open,high,low,close,volume', "$Date,100,101,99,100,1000"))
}
try {
    # Parse every production script in the supplied snapshot before running anything.
    foreach ($file in Get-ChildItem "$sourceRoot/ps" -Filter '*.ps1') {
        $tokens=$null; $errors=$null
        [void][Management.Automation.Language.Parser]::ParseFile($file.FullName,[ref]$tokens,[ref]$errors)
        Assert ($errors.Count -eq 0) "Parse failure $($file.Name): $errors"
    }
    Assert ((Get-CompletedTradingDate -NowJapan ([datetime]'2026-11-11T08:00')).ToString('yyyy-MM-dd') -eq $script:target) 'Morning target changed'
    Assert ((Get-CompletedTradingDate -NowJapan ([datetime]'2026-11-11T10:00')).ToString('yyyy-MM-dd') -eq $script:target) 'Delayed morning used live session'
    Assert ((Get-CompletedTradingDate -NowJapan ([datetime]'2026-11-11T16:59')).ToString('yyyy-MM-dd') -eq $script:target) 'Pre-17:00 session accepted'
    Assert ((Get-CompletedTradingDate -NowJapan ([datetime]'2026-11-11T17:00')).ToString('yyyy-MM-dd') -eq '2026-11-11') 'Evening target wrong'
    Assert ((Get-CompletedTradingDate -NowJapan ([datetime]'2026-11-09T08:00')).ToString('yyyy-MM-dd') -eq '2026-11-06') 'Weekend boundary wrong'
    Assert ((Get-CompletedTradingDate -NowJapan ([datetime]'2026-11-03T18:00')).ToString('yyyy-MM-dd') -eq '2026-11-02') 'Holiday boundary wrong'
    $rejected=$false
    try { Resolve-DailyTarget -AsOf '2026-11-11' -NowJapan ([datetime]'2026-11-11T10:00') | Out-Null } catch { $rejected=$true }
    Assert $rejected 'Unclosed target accepted'
    Assert ((Get-MarketDateLimit -AsOf $script:target -Offset -5) -eq '2026-11-09') 'US unclosed daily bar allowed'
    Assert ((Get-MarketDateLimit -AsOf $script:target -Offset 1) -eq '2026-11-09') 'FX unclosed daily bar allowed'
    Assert (-not (Test-MarketLastDate -Latest '2026-10-10' -AsOf $script:target -Offset -5)) 'Month-old non-JP history accepted'
    Assert (Test-MarketLastDate -Latest '2026-11-06' -AsOf $script:target -Offset -5) 'Normal non-JP holiday gap rejected'
    $rows=@(ConvertFrom-MarketChart -Response (New-ChartResponse $true) -Offset 9 -LastDate $script:target)
    Assert ($rows.Count -eq 1 -and $rows[0].date -eq '2026-11-09') 'Null target or future market bar leaked'
    $bad=New-ChartResponse $false; $bad.chart.result[0].indicators.quote[0].open[1]=[double]::PositiveInfinity
    Assert (@(ConvertFrom-MarketChart -Response $bad -Offset 9 -LastDate $script:target).Count -eq 1) 'Infinite quote accepted'

    New-Item -ItemType Directory -Force "$fixture/ps", "$fixture/data/raw/market/daily_since2000", "$fixture/data/raw/stocks/prime_since2000", "$fixture/reports" | Out-Null
    foreach ($name in @('Common.ps1','DailyData.ps1','Fetch-MarketData.ps1','Fetch-UniversePrices.ps1','Test-DailyUpdate.ps1','Test-NeedsUpdate.ps1')) {
        Copy-Item "$sourceRoot/ps/$name" "$fixture/ps/$name"
    }
    # HTTP 200 + null OHLC on the target must retry, then recover, with no later bar saved.
    & "$fixture/ps/Fetch-MarketData.ps1" -AsOf $script:target -RetryDelaySeconds 1 -MaxAttempts 3
    Assert ($script:requestCounts['^N225'] -eq 2) 'HTTP-success missing target was not retried'
    Assert ($script:sleeps.Contains(1)) 'Missing target retry did not wait'
    Assert ((Get-DailyLastDate "$fixture/data/raw/market/daily_since2000/N225.csv") -eq $script:target) 'Recovered target not saved'
    $md=@(Get-Content "$fixture/logs/market_fetch.json" -Raw | ConvertFrom-Json)
    $nk=@($md | Where-Object symbol -eq '^N225')
    Assert ($nk.Count -eq 2 -and -not $nk[0].ok -and $nk[1].ok) 'Retry diagnostics wrong'
    Assert ($nk[0].latestReturnedDate -eq '2026-11-11' -and $nk[0].latestValidDate -eq '2026-11-09') 'Null/future diagnostic wrong'

    # Persistent target miss fails after a bounded number of attempts and preserves old CSV.
    $script:mode='market-failure'; $script:requestCounts=@{}
    Write-PriceFile "$fixture/data/raw/market/daily_since2000/N225.csv" '2026-11-06'
    $before=[IO.File]::ReadAllText("$fixture/data/raw/market/daily_since2000/N225.csv")
    $failed=$false
    try { & "$fixture/ps/Fetch-MarketData.ps1" -AsOf $script:target -RetryDelaySeconds 1 -MaxAttempts 3 } catch { $failed=$true }
    Assert $failed 'Permanent target miss succeeded'
    Assert ($script:requestCounts['^N225'] -eq 3) 'Retry limit ignored'
    Assert ([IO.File]::ReadAllText("$fixture/data/raw/market/daily_since2000/N225.csv") -ceq $before) 'Incomplete fetch overwrote old CSV'

    # Stock getter must enforce EndDate even if the server returns later rows.
    $script:mode='api-error'; $script:requestCounts=@{}
    $stockRows=@(Get-YahooChart -Code '3333' -Years 2 -WithRawPrices -EndDate $script:target)
    Assert ($script:requestCounts['3333.T'] -eq 2 -and $stockRows[-1].date -eq $script:target) 'Transport retry or stock date cap failed'
    $script:mode='stocks'; $script:requestCounts=@{}
    [IO.File]::WriteAllLines("$fixture/universe.csv", @('code,name','1111,Recover','2222,Suspended'))
    @{universe=@{csv='universe.csv'};data=@{historyYears=2;startDate='2020-01-01';tickerDir='data/raw/stocks/prime_since2000'}} | ConvertTo-Json -Depth 4 | Set-Content "$fixture/ps/config.json" -Encoding UTF8
    $env:KABU_CONFIG=''
    Write-PriceFile "$fixture/data/raw/stocks/prime_since2000/2222.csv" '2026-11-06'
    $before=[IO.File]::ReadAllText("$fixture/data/raw/stocks/prime_since2000/2222.csv")
    & "$fixture/ps/Fetch-UniversePrices.ps1" -Refresh -AsOf $script:target -RetryDelaySeconds 1 -MaxAttempts 3
    Assert ($script:requestCounts['1111.T'] -eq 2 -and $script:requestCounts['2222.T'] -eq 3) 'Selective stock retry failed'
    Assert ((Get-DailyLastDate "$fixture/data/raw/stocks/prime_since2000/1111.csv") -eq $script:target) 'Recovered stock missing'
    Assert ([IO.File]::ReadAllText("$fixture/data/raw/stocks/prime_since2000/2222.csv") -ceq $before) 'Stale stock overwrote cache'

    # Exercise the real pre-aggregation gate in a child process on synthetic CSVs only.
    Get-ChildItem "$fixture/data/raw/stocks/prime_since2000" -Filter '*.csv' | Remove-Item
    foreach ($name in @('1321','N225','1571','1357','2558','2559')) { Write-PriceFile "$fixture/data/raw/market/daily_since2000/$name.csv" $script:target }
    foreach ($name in @('GSPC','VIX','USDJPY')) { Write-PriceFile "$fixture/data/raw/market/daily_since2000/$name.csv" '2026-11-09' }
    foreach ($n in 1..100) { Write-PriceFile "$fixture/data/raw/stocks/prime_since2000/$n.csv" $script:target }
    New-Item -ItemType Directory -Force "$fixture/data/raw/universe" | Out-Null
    $universeLines = @('code,name') + @(1..100 | ForEach-Object { "$_,Synthetic" })
    [IO.File]::WriteAllLines("$fixture/data/raw/universe/prime.csv", $universeLines)
    $exe=Get-PowerShellExe
    & $exe -NoProfile -File "$fixture/ps/Test-DailyUpdate.ps1" -NoRepair -PricesOnly -AsOf $script:target *> "$fixture/gate-good.log"
    Assert ($LASTEXITCODE -eq 0) 'Complete synthetic prices failed gate'
    foreach ($n in 1..3) { Write-PriceFile "$fixture/data/raw/stocks/prime_since2000/$n.csv" '2026-11-09' }
    & $exe -NoProfile -File "$fixture/ps/Test-DailyUpdate.ps1" -NoRepair -PricesOnly -AsOf $script:target *> "$fixture/gate-stale.log"
    Assert ($LASTEXITCODE -ne 0) 'Fresh mtime concealed stale content'
    foreach ($n in 1..3) { Write-PriceFile "$fixture/data/raw/stocks/prime_since2000/$n.csv" $script:target }
    Write-PriceFile "$fixture/data/raw/stocks/prime_since2000/1.csv" '2026-11-11'
    & $exe -NoProfile -File "$fixture/ps/Test-DailyUpdate.ps1" -NoRepair -PricesOnly -AsOf $script:target *> "$fixture/gate-future.log"
    Assert ($LASTEXITCODE -ne 0) 'One future cached row escaped gate'
    Write-PriceFile "$fixture/data/raw/stocks/prime_since2000/1.csv" $script:target
    Write-PriceFile "$fixture/data/raw/market/daily_since2000/N225.csv" '2026-11-09'
    & $exe -NoProfile -File "$fixture/ps/Test-DailyUpdate.ps1" -NoRepair -PricesOnly -AsOf $script:target *> "$fixture/gate-index.log"
    Assert ($LASTEXITCODE -ne 0) 'Missing N225 target escaped gate'
    Write-PriceFile "$fixture/data/raw/market/daily_since2000/N225.csv" $script:target
    foreach ($n in 1..3) { Remove-Item "$fixture/data/raw/stocks/prime_since2000/$n.csv" }
    & $exe -NoProfile -File "$fixture/ps/Test-DailyUpdate.ps1" -NoRepair -PricesOnly -AsOf $script:target *> "$fixture/gate-missing.log"
    Assert ($LASTEXITCODE -ne 0) 'Missing expected CSVs escaped denominator'
    foreach ($n in 1..3) {
        Write-PriceFile "$fixture/data/raw/stocks/prime_since2000/$n.csv" $script:target
        $path="$fixture/data/raw/stocks/prime_since2000/$n.csv"
        [IO.File]::WriteAllText($path, [IO.File]::ReadAllText($path).Replace(',100,101,99,100,1000', ',NaN,101,99,100,1000'))
    }
    & $exe -NoProfile -File "$fixture/ps/Test-DailyUpdate.ps1" -NoRepair -PricesOnly -AsOf $script:target *> "$fixture/gate-invalid.log"
    Assert ($LASTEXITCODE -ne 0) 'Invalid target OHLC escaped content validation'
    Write-Output 'PASS: retry on missing target, bounded failures, cache preservation, API error retry, stock/market cutoff, JST/holiday boundaries, and pre-aggregation freshness gate'
    Write-Output "Synthetic evidence: $fixture"
} finally {
    $env:KABU_CONFIG=$oldConfig; $env:GITHUB_OUTPUT=$oldOutput; Set-Location $oldLocation
}
