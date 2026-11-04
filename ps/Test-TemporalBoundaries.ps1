# Offline synthetic regression; never downloads quotes or uses portfolio data.
param([switch]$SkipModel)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
. "$PSScriptRoot/Common.ps1"
. "$PSScriptRoot/RealizationTiming.ps1"
Add-Type -Path "$PSScriptRoot/CrossSectionStudy.cs"
function Assert($ok, $message) { if (-not $ok) { throw $message } }
$fixture = Join-Path ([IO.Path]::GetTempPath()) ('kabu-temporal-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path "$fixture/prices", "$fixture/ps", "$fixture/data/raw/market/daily_since2000" -Force | Out-Null
$dates = @(); $d = [datetime]'2023-01-02'
while ($dates.Count -lt 490) {
    if ($d.DayOfWeek -notin @('Saturday', 'Sunday') -and $d.DayOfYear % 61 -ne 0) { $dates += $d.ToString('yyyy-MM-dd') }
    $d = $d.AddDays(1)
}
$header = 'date,code,open,high,low,close,volume,raw_open,raw_high,raw_low,raw_close'
$flat = @($header); foreach ($date in $dates) { $flat += "$date,A,100,101,99,100,10000000,100,101,99,100" }
[IO.File]::WriteAllLines("$fixture/prices/A.csv", $flat)
[IO.File]::WriteAllLines("$fixture/prices/B.csv", $flat.Replace(',A,', ',B,'))
[IO.File]::WriteAllLines("$fixture/calendar.csv", $flat)
$before = [CrossSectionStudy]::Run("$fixture/prices", "$fixture/calendar.csv", 1, 1, 1, "$fixture/invalid.log", 0)
$baseRow = @($before | Where-Object Date -eq $dates[260])[0]
Assert ($baseRow.Picks[8] -eq 'A') 'Synthetic fixture must select A'
$flat[262] = $flat[262].Replace('10000000', '0')
[IO.File]::WriteAllLines("$fixture/prices/A.csv", $flat)
$after = [CrossSectionStudy]::Run("$fixture/prices", "$fixture/calendar.csv", 1, 1, 1, "$fixture/invalid.log", 0)
$changed = @($after | Where-Object Date -eq $dates[260])[0]
Assert ($changed.Picks[8] -eq $baseRow.Picks[8]) 'Future availability changed selection'
Assert ([double]::IsNaN($changed.Signal[8])) 'Failed execution must remain missing'
Assert ($changed.MissingOutcomes -eq 1 -and -not [double]::IsNaN($changed.MarketEw)) 'Missing outcome poisoned market mean'
$rows = @([pscustomobject]@{ sell_date='2024-01-05'; target_date='2024-01-04'; rv_pct2=1 }, [pscustomobject]@{ sell_date='2024-01-09'; target_date='2024-01-05'; rv_pct2=999 })
Assert ((Get-RealizedLags $rows '2024-01-05')[0] -eq 1) 'Future label entered lag'
$rows[1].rv_pct2 = 2000
Assert ((Get-RealizedLags $rows '2024-01-05')[0] -eq 1) 'Future perturbation changed lag'
Assert ([double]::IsNaN((Get-RealizedLags $rows '2024-01-05')[1])) 'Insufficient history must be missing'
Assert ((Resolve-ProjectPath "$fixture/absolute.csv") -eq "$fixture/absolute.csv") 'Absolute data path was rebased'
# Load only the production valuation function, without running the data pipeline.
$tokens=$null; $errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile("$PSScriptRoot/Simulate-SKabu.ps1", [ref]$tokens, [ref]$errors)
Assert ($errors.Count -eq 0) 'Simulation script parse error'
$fn=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-LastKnownPositionValue'}, $true)
. ([scriptblock]::Create($fn.Extent.Text))
$px=@{A=[pscustomobject]@{idx=@{'2024-01-01'=0;'2024-01-03'=1;'2024-01-08'=2};adjClose=@(100,150,900)}}
$position=[pscustomobject]@{code='A';shares=2;entryRaw=100;entryAdj=100}
Assert ((Get-LastKnownPositionValue $position '2024-01-05') -eq 300) 'Missing-day mark used basis or future value'
Assert ((Get-LastKnownPositionValue $position '2023-12-31') -eq 200) 'No-history mark must use entry basis'
if (-not $SkipModel) {
    foreach ($name in @('Common.ps1','RealizationTiming.ps1','Build-OvernightVolModel.ps1')) { Copy-Item -LiteralPath "$PSScriptRoot/$name" -Destination "$fixture/ps/$name" }
    $random=[Random]::new(1234)
    foreach ($code in @('A','B','N225')) {
        $lines=@($header); $price=100.0
        foreach ($date in $dates) {
            $open=$price*(0.99+0.02*$random.NextDouble()); $close=$open*(0.99+0.02*$random.NextDouble())
            $hi=[Math]::Max($open,$close)*1.01; $lo=[Math]::Min($open,$close)*0.99
            $lines += [string]::Format([Globalization.CultureInfo]::InvariantCulture,'{0},{1},{2},{3},{4},{5},10000000,{2},{3},{4},{5}', $date,$code,$open,$hi,$lo,$close)
            $price=$close
        }
        if ($code -eq 'N225') { [IO.File]::WriteAllLines("$fixture/data/raw/market/daily_since2000/N225.csv", $lines) }
        else {
            # One outcome is unavailable, with earlier history still intact.
            if ($code -eq 'A') { $lines = @($lines | Where-Object { -not $_.StartsWith($dates[420]+',') }) }
            [IO.File]::WriteAllLines("$fixture/prices/$code.csv", $lines)
        }
    }
    $ranking=@('date,rank,code'); foreach ($date in $dates) { $ranking += "$date,1,A"; $ranking += "$date,2,B" }
    [IO.File]::WriteAllLines("$fixture/ranking.csv", $ranking)
    @{ universe=@{topNPopular=2}; data=@{dailyRankingCsv='ranking.csv'}; skabu=@{tickerDir='prices'} } | ConvertTo-Json | Set-Content -LiteralPath "$fixture/ps/config.json" -Encoding utf8
    $oldConfig=$env:KABU_CONFIG; $oldLocation=Get-Location
    try {
        $env:KABU_CONFIG="$fixture/ps/config.json"
        & "$fixture/ps/Build-OvernightVolModel.ps1" -FirstPredictYear 2024 -MinTrain 50 -OutDir 'results' *> "$fixture/build.log"
    } finally { $env:KABU_CONFIG=$oldConfig; Set-Location $oldLocation }
    $dataset=Import-Csv "$fixture/results/dataset.csv"
    foreach ($row in $dataset) {
        $lags=Get-RealizedLags $dataset $row.forecast_date
        foreach ($i in 0..2) {
            $actual=[double]$row.(@('rv_lag1','rv_lag5','rv_lag22')[$i])
            Assert (([double]::IsNaN($actual) -and [double]::IsNaN($lags[$i])) -or [Math]::Abs($actual-$lags[$i]) -lt 1e-10) 'Dataset lag exceeds realization cutoff'
        }
    }
    $models=@(Get-ChildItem "$fixture/results/model_*.json")
    Assert ($models.Count -gt 0) 'No model was trained'
    foreach ($file in $models) { $model=Get-Content $file -Raw | ConvertFrom-Json; Assert ($model.train_realization_max -le $model.training_cutoff) 'Future labels entered training' }
    $forecasts=@(Import-Csv "$fixture/results/forecasts.csv")
    Assert ($forecasts.Count -gt 0) 'No forecasts'
    Assert (@($forecasts[0].PSObject.Properties.Name | Where-Object { $_ -match 'realized' }).Count -eq 0) 'Forecast output includes outcome labels'
    $predictions=@(Import-Csv "$fixture/results/predictions.csv")
    Assert (@($predictions | Where-Object { [double]::IsNaN([double]$_.realized_variance_pct2) }).Count -gt 0) 'Missing outcomes blocked forecasts'
}
Write-Output 'PASS: selection, missing outcomes, lag cutoff, absolute paths, delayed valuation, training cutoff and forecast separation'
Write-Output "Synthetic evidence: $fixture"
