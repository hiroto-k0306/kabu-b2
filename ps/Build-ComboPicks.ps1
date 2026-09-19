# 売買代金ランキングと B2 の選定結果から、比較用ユニバース(combo)の date,rank,code を作る。
#   nikkei_top10.csv : 日経225(各時点構成)の売買代金 上位10
#   prime_top10.csv  : 東証プライムの売買代金 上位10
#   mix_each5.csv    : 日経225 上位5(rank1-5) + B2 上位5(rank6-10)
# 先に Build-TurnoverRanking.ps1 (nikkei_pit / prime) と Export-B3LPicks.ps1 を実行しておく。
# 使い方: powershell -File ps\Build-ComboPicks.ps1
param(
    [string]$NikkeiRanking = "data/processed/nikkei_pit/daily_turnover_ranking_5y.csv",
    [string]$PrimeRanking  = "data/processed/prime/daily_turnover_ranking_5y.csv",
    [string]$B3LTop5       = "data/processed/b3l/daily_picks_top5.csv",
    [string]$OutDir        = "data/processed/combo"
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"
Set-Location (Get-ProjectRoot)

function Import-RankedCsv {
    # date -> rank順に並べたcodeの配列
    param([string]$Path)
    $full = Resolve-ProjectPath $Path
    if (-not (Test-Path $full)) { throw "$Path がない" }
    $map = @{}
    foreach ($r in (Import-Csv -Path $full -Encoding UTF8)) {
        if (-not $map.ContainsKey($r.date)) { $map[$r.date] = New-Object System.Collections.Generic.List[object] }
        $map[$r.date].Add([PSCustomObject]@{ rank = [int]$r.rank; code = [string]$r.code })
    }
    $out = @{}
    foreach ($d in $map.Keys) { $out[$d] = @($map[$d] | Sort-Object rank | Select-Object -ExpandProperty code) }
    return $out
}

$nk    = Import-RankedCsv $NikkeiRanking
$prime = Import-RankedCsv $PrimeRanking
$b3l   = Import-RankedCsv $B3LTop5

$outDirFull = Split-Path (Resolve-ProjectPath "$OutDir/placeholder") -Parent
New-Item -ItemType Directory -Force $outDirFull | Out-Null

function Export-Picks {
    param([string]$Name, [object[]]$Rows)
    $path = Join-Path $outDirFull $Name
    $Rows | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
    $days = ($Rows | Select-Object -ExpandProperty date -Unique).Count
    Write-Host ("{0,-18} {1,6} rows / {2,5} days  {3} .. {4}" -f $Name, $Rows.Count, $days, $Rows[0].date, $Rows[-1].date)
}

# 上位N をそのまま写す
foreach ($t in @(@{ src = $nk; name = "nikkei_top10.csv" }, @{ src = $prime; name = "prime_top10.csv" })) {
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($d in ($t.src.Keys | Sort-Object)) {
        $codes = $t.src[$d]
        if ($codes.Count -lt 10) { continue }
        for ($k = 0; $k -lt 10; $k++) { $rows.Add([PSCustomObject]@{ date = $d; rank = $k + 1; code = $codes[$k] }) }
    }
    Export-Picks -Name $t.name -Rows $rows.ToArray()
}

# 日経上位5 と B2 上位5 を半分ずつ。両方そろう日だけ
$rows = New-Object System.Collections.Generic.List[object]
foreach ($d in ($nk.Keys | Sort-Object)) {
    if (-not $b3l.ContainsKey($d)) { continue }
    $a = $nk[$d]; $b = $b3l[$d]
    if ($a.Count -lt 5 -or $b.Count -lt 5) { continue }
    for ($k = 0; $k -lt 5; $k++) { $rows.Add([PSCustomObject]@{ date = $d; rank = $k + 1; code = $a[$k] }) }
    for ($k = 0; $k -lt 5; $k++) { $rows.Add([PSCustomObject]@{ date = $d; rank = $k + 6; code = $b[$k] }) }
}
Export-Picks -Name "mix_each5.csv" -Rows $rows.ToArray()
