# Test-CrossSectionSignals.ps1 の出力(picks_B3L)を、Simulate-SKabu が読む date,rank,code 形式へ変換する。
# picks_B3L は TopK 個のコードを半角空白で連結した文字列。TopK=10 で出せば top5/top3/top1 はその先頭を切るだけでよい。
# 使い方: powershell -File ps\Export-B3LPicks.ps1
param(
    [string]$SourceCsv = "reports/cross_section/daily_Reference_minprice0.csv",
    [string]$OutDir    = "data/processed/b3l",
    [string]$Column    = "picks_B3L"
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"
Set-Location (Get-ProjectRoot)

$src = Resolve-ProjectPath $SourceCsv
if (-not (Test-Path $src)) { throw "$SourceCsv がない。先に Test-CrossSectionSignals.ps1 -Phase Reference -Signal B3L -TopK 10 を実行する" }
$rows = Import-Csv -Path $src -Encoding UTF8
if (-not ($rows[0].PSObject.Properties.Name -contains $Column)) { throw "$SourceCsv に $Column 列がない" }

$outDirFull = Split-Path (Resolve-ProjectPath "$OutDir/placeholder") -Parent
New-Item -ItemType Directory -Force $outDirFull | Out-Null

# 出力先 -> 採用する上位何銘柄か
$targets = [ordered]@{ "daily_picks.csv" = 10; "daily_picks_top5.csv" = 5; "daily_picks_top3.csv" = 3; "daily_picks_top1.csv" = 1 }

foreach ($name in $targets.Keys) {
    $topN = $targets[$name]
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($r in $rows) {
        $picks = [string]$r.$Column
        if ([string]::IsNullOrWhiteSpace($picks)) { continue }   # 対象銘柄が足りない日は空になる
        $codes = @($picks -split '\s+' | Where-Object { $_ -ne "" })
        if ($codes.Count -lt $topN) { continue }
        for ($k = 0; $k -lt $topN; $k++) {
            $out.Add([PSCustomObject]@{ date = $r.date; rank = $k + 1; code = $codes[$k] })
        }
    }
    $path = Join-Path $outDirFull $name
    $out | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
    $days = ($out | Select-Object -ExpandProperty date -Unique).Count
    Write-Host ("{0,-22} {1,6} rows / {2,5} days  {3} .. {4}" -f $name, $out.Count, $days, $out[0].date, $out[-1].date)
}
