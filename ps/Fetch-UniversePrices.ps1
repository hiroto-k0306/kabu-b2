# 銘柄数の多いユニバース(東証プライム等)向けの株価取得。
# 銘柄ごとに data.tickerDir/<code>.csv へ保存するので、途中で止まっても再実行すれば続きから取得する。
# 1ファイルに調整済み価格(open/high/low/close/volume)と、分割・配当調整前の実際の価格(raw_*)の両方を持つ。
# 使い方: $env:KABU_CONFIG = "ps/config.prime.json"; powershell -File ps\Fetch-UniversePrices.ps1

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"

$config = Get-Config
$universe = Import-Csv -Path (Resolve-ProjectPath $config.universe.csv) -Encoding UTF8
$tickerDir = Split-Path (Resolve-ProjectPath "$($config.data.tickerDir)/placeholder") -Parent

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$fetched = 0; $skipped = 0
$failed = New-Object System.Collections.Generic.List[string]
$i = 0
foreach ($u in $universe) {
    $i++
    $path = Join-Path $tickerDir "$($u.code).csv"
    if ((Test-Path $path) -and (Get-Item $path).Length -gt 0) { $skipped++; continue }

    $startDate = ""
    if ($config.data.startDate) { $startDate = [string]$config.data.startDate }
    $rows = Get-YahooChart -Code $u.code -Years $config.data.historyYears -WithRawPrices -StartDate $startDate
    if ($rows.Count -eq 0) {
        $failed.Add($u.code)
        Write-Warning "[$i/$($universe.Count)] $($u.code) $($u.name): no data"
        Start-Sleep -Milliseconds 500
        continue
    }
    $tmp = "$path.tmp"
    $rows | Export-Csv -Path $tmp -NoTypeInformation -Encoding UTF8
    Move-Item -Path $tmp -Destination $path -Force
    $fetched++
    if ($i % 50 -eq 0) { Write-Host ("[{0}/{1}] fetched {2}, skipped {3}, failed {4}, elapsed {5}" -f $i, $universe.Count, $fetched, $skipped, $failed.Count, $sw.Elapsed) }
    Start-Sleep -Milliseconds 200
}

Write-Host ("done: fetched {0}, already had {1}, failed {2}, elapsed {3}" -f $fetched, $skipped, $failed.Count, $sw.Elapsed)
if ($failed.Count -gt 0) {
    Write-Host "failed codes: $($failed -join ', ')"
    $failed | Set-Content -Path (Join-Path $tickerDir "_failed.txt") -Encoding UTF8
}
