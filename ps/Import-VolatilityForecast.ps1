# 別プロジェクト(kabu)の変動率モデルが出した「翌営業日の予測変動率」を取り込む。
#
# あちらのモデルは日本時間16:00に翌営業日の Parkinson分散(%^2)を予測する。
# B2 は前営業日の引けまでで銘柄を決め、当日10:30〜14:00に注文して引けで買う。
# したがって買付日 D に使えるのは「D-1 の16:00に作られた、D を対象とする予測」で、
# これは B2 の選定に使う情報と同じ締め切り(D-1の引け)に収まる。
#
# 取り込み元(既定 ../kabu):
#   reports/nikkei_volatility_predictions.csv  最終評価649日ぶん(対象 2024-01-04〜2026-08-31)
#   reports/paper/evaluation.csv               予測記録(対象 2026-09-01〜)
# 出力: data/processed/vol/nikkei_vol_forecast.csv と source_meta.json
#
# 使い方: powershell -File ps\Import-VolatilityForecast.ps1
param(
    [string]$KabuRoot   = "../kabu",
    [string]$HoldoutCsv = "reports/nikkei_volatility_predictions.csv",
    [string]$PaperCsv   = "reports/paper/evaluation.csv",
    [string]$OutCsv     = "data/processed/vol/nikkei_vol_forecast.csv"
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"
Set-Location (Get-ProjectRoot)
$root = Get-ProjectRoot

$kabu = $KabuRoot
if (-not [IO.Path]::IsPathRooted($kabu)) { $kabu = Join-Path $root $KabuRoot }
if (-not (Test-Path $kabu)) { throw "取り込み元が見つからない: $kabu" }
$kabu = (Resolve-Path $kabu).Path

# target_date -> 行。あとから読む方を優先しないので、重複は値が違うときだけ知らせる
$rows = New-Object 'System.Collections.Generic.Dictionary[string,object]'
$skipped = 0

function Add-Forecast {
    param([string]$Path, [string]$Source)
    $full = Join-Path $kabu $Path
    if (-not (Test-Path $full)) { Write-Warning "$Source : $Path がない。飛ばす"; return 0 }
    $n = 0
    foreach ($r in (Import-Csv -Path $full -Encoding UTF8)) {
        $tgt = [string]$r.target_date
        $fc  = [string]$r.forecast_date
        if (-not $tgt -or -not $fc) { $script:skipped++; continue }
        # 予測日が対象日より後ろなら未来の情報が混じっている
        if ([string]::Compare($fc, $tgt) -ge 0) { Write-Warning "未来情報の疑い: forecast=$fc target=$tgt。飛ばす"; $script:skipped++; continue }
        $sigma = 0.0
        if (-not [double]::TryParse([string]$r.predicted_pk_sigma_pct, [ref]$sigma)) { $script:skipped++; continue }
        if (-not ($sigma -gt 0)) { $script:skipped++; continue }
        $var = 0.0
        [void][double]::TryParse([string]$r.predicted_pk_variance_pct2, [ref]$var)

        if ($rows.ContainsKey($tgt)) {
            $old = $rows[$tgt]
            if ([Math]::Abs($old.pred_sigma_pct - $sigma) -gt 1e-9) {
                Write-Warning "$tgt の予測が元ファイル間で食い違う($($old.source)=$($old.pred_sigma_pct) / $Source=$sigma)。先に読んだ方を残す"
            }
            continue
        }
        $rows[$tgt] = [PSCustomObject]@{
            target_date       = $tgt
            forecast_date     = $fc
            pred_sigma_pct    = $sigma
            pred_variance_pct2 = $var
            source            = $Source
        }
        $n++
    }
    Write-Host ("{0,-22} {1,5}件  {2}" -f $Source, $n, $Path)
    return $n
}

[void](Add-Forecast -Path $HoldoutCsv -Source "holdout")
[void](Add-Forecast -Path $PaperCsv   -Source "paper")

if ($rows.Count -eq 0) { throw "取り込める予測が1件もない" }

$sorted = @($rows.Values | Sort-Object target_date)
$outPath = Resolve-ProjectPath $OutCsv
New-Item -ItemType Directory -Force (Split-Path $outPath -Parent) | Out-Null
$sorted | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8

# 営業日カレンダー(N225)と突き合わせて、抜けと余りを知らせる
$cal = New-Object System.Collections.Generic.HashSet[string]
$n225 = Resolve-ProjectPath "data/raw/market/daily_since2000/N225.csv"
if (Test-Path $n225) {
    foreach ($r in (Import-Csv -Path $n225 -Encoding UTF8)) { [void]$cal.Add([string]$r.date) }
}
$first = $sorted[0].target_date; $last = $sorted[-1].target_date
$inRange = @($cal | Where-Object { [string]::Compare($_, $first) -ge 0 -and [string]::Compare($_, $last) -le 0 })
$have = New-Object System.Collections.Generic.HashSet[string]
foreach ($r in $sorted) { [void]$have.Add($r.target_date) }
$missing = @($inRange | Where-Object { -not $have.Contains($_) } | Sort-Object)
$extra   = @($sorted | Where-Object { -not $cal.Contains($_.target_date) } | Select-Object -ExpandProperty target_date)

$sigmas = @($sorted | Select-Object -ExpandProperty pred_sigma_pct | Sort-Object)
$median = $sigmas[[int][Math]::Floor($sigmas.Count / 2)]

Write-Host ""
Write-Host ("保存 {0}: {1}件  対象 {2} .. {3}" -f $OutCsv, $sorted.Count, $first, $last)
Write-Host ("  予測σ(%): 中央値 {0:N3}  最小 {1:N3}  最大 {2:N3}" -f $median, $sigmas[0], $sigmas[-1])
Write-Host ("  N225営業日で予測が無い日: {0}件{1}" -f $missing.Count, $(if ($missing.Count -gt 0) { " 例 " + (($missing | Select-Object -First 5) -join ", ") } else { "" }))
Write-Host ("  N225営業日でない対象日  : {0}件{1}" -f $extra.Count, $(if ($extra.Count -gt 0) { " " + ($extra -join ", ") + " (先の営業日ぶん)" } else { "" }))
if ($skipped -gt 0) { Write-Host "  検査で落とした行: $skipped" }

# 元ファイルのハッシュを残す(あちらの流儀にそろえる)
$meta = [ordered]@{
    imported_at = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    kabu_root   = $kabu
    rows        = $sorted.Count
    target_from = $first
    target_to   = $last
    median_sigma_pct = $median
    missing_trading_days = $missing.Count
    sources     = [ordered]@{}
}
foreach ($pair in @(@{ p = $HoldoutCsv; n = "holdout" }, @{ p = $PaperCsv; n = "paper" })) {
    $full = Join-Path $kabu $pair.p
    if (Test-Path $full) {
        $meta.sources[$pair.n] = [ordered]@{
            path   = $pair.p
            sha256 = (Get-FileHash -Path $full -Algorithm SHA256).Hash.ToLower()
        }
    }
}
$metaPath = Join-Path (Split-Path $outPath -Parent) "source_meta.json"
[IO.File]::WriteAllText($metaPath, ([PSCustomObject]$meta | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding $false))
Write-Host "saved $(Split-Path $OutCsv -Parent)/source_meta.json"
