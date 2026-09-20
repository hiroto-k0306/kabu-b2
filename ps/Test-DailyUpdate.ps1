# 17時の日次更新がちゃんと通ったかを確かめ、駄目なら取り直す。
# タスクスケジューラから翌朝8時に呼ばれる想定。
# 朝に動くので、見るのは「今日」ではなく直近の営業日(ふつうは前営業日)。
# 使い方: powershell -File ps\Test-DailyUpdate.ps1
param(
    [switch]$Force,        # 取引時間中でも取り直す(ふつうは使わない)
    [string]$AsOf = "",    # この日付(yyyy-MM-dd)のデータを見る。既定は直近の営業日
    [switch]$NoRepair,     # 異常でも取り直さず、記録だけする
    [int]$Budget = 500000,
    [double]$MinFreshRatio = 0.98,   # 当日付けで更新されていてほしい銘柄の割合
    [int]$SampleSize = 30,           # 中身まで見る銘柄数
    [string]$LogDir = "logs"
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"
Set-Location (Get-ProjectRoot)
$root = Get-ProjectRoot

# どの日のデータを見るか。朝に動かす想定なので既定は直近の営業日。
# 夕方(17時)以降に手で動かしたときだけ当日を見る。
$now = Get-Date
if ($AsOf) {
    $target = [datetime]::ParseExact($AsOf, "yyyy-MM-dd", $null)
} elseif ((Test-TradingDay -Date $now.Date) -and $now.Hour -ge 17) {
    $target = $now.Date
} else {
    $target = Get-PrevTradingDay -Date $now.Date
}
$targetStr = $target.ToString("yyyy-MM-dd")
$logDirFull = Join-Path $root $LogDir
New-Item -ItemType Directory -Force $logDirFull | Out-Null
$logPath    = Join-Path $logDirFull ("check_{0}.log" -f $now.ToString("yyyyMMdd"))
$statusPath = Join-Path $logDirFull "last_check.json"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "{0} [{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Write-Host $line
    Add-Content -Path $logPath -Value $line -Encoding UTF8
}

function Get-LastDate {
    # CSVの最終行の日付("yyyy-MM-dd")。読めなければ空文字
    param([string]$Path)
    try {
        $last = Get-Content -Path $Path -Tail 1 -ErrorAction Stop
        if (-not $last) { return "" }
        return ($last -split ",")[0].Trim('"')
    } catch { return "" }
}

function Test-Freshness {
    <# 当日のデータがそろっているかを見る。理由の一覧を返す(空なら正常) #>
    $problems = New-Object System.Collections.Generic.List[string]

    # 1. 日経平均(営業日カレンダーの元)が当日まで来ているか
    $n225 = Join-Path $root "data\raw\market\daily_since2000\N225.csv"
    $n225Last = Get-LastDate $n225
    if ($n225Last -ne $targetStr) { $problems.Add("N225.csv の最終日が $n225Last (期待 $targetStr)") }
    else { Write-Log "  N225.csv: $n225Last  OK" }

    # 2. 銘柄CSVが当日に書き直されているか(更新時刻で素早く見る)
    $tickerDir = Join-Path $root "data\raw\stocks\prime_since2000"
    if (-not (Test-Path $tickerDir)) {
        $problems.Add("$tickerDir がない")
        return $problems
    }
    $files = @(Get-ChildItem -Path $tickerDir -Filter "*.csv" -File)
    $fresh = @($files | Where-Object { $_.LastWriteTime -ge $target }).Count
    $ratio = if ($files.Count -gt 0) { $fresh / $files.Count } else { 0 }
    Write-Log ("  銘柄CSV: {0}/{1} が $targetStr 以降に更新 ({2:P1})" -f $fresh, $files.Count, $ratio)
    if ($ratio -lt $MinFreshRatio) { $problems.Add(("$targetStr 以降に更新された銘柄CSVが {0}/{1} しかない" -f $fresh, $files.Count)) }

    # 3. 中身も当日の行で終わっているか(全部読むと重いので抜き取り)
    $sample = @($files | Get-Random -Count ([Math]::Min($SampleSize, $files.Count)))
    $stale = @($sample | Where-Object { (Get-LastDate $_.FullName) -ne $targetStr })
    Write-Log ("  抜き取り {0}銘柄: 最終行が $targetStr でないもの {1}件" -f $sample.Count, $stale.Count)
    # 上場廃止・取引停止の銘柄は当日の行を持たないので、少数なら異常としない
    if ($stale.Count -gt [Math]::Max(2, $sample.Count * 0.1)) {
        $problems.Add(("抜き取り{0}銘柄中{1}銘柄の最終行が $targetStr でない: {2}" -f $sample.Count, $stale.Count, (($stale | Select-Object -First 5 | ForEach-Object { $_.BaseName }) -join ", ")))
    }

    # 4. 翌営業日の銘柄が当日のデータで作られているか
    $picks = Join-Path $root "reports\today_picks.json"
    if (-not (Test-Path $picks)) { $problems.Add("reports/today_picks.json がない") }
    else {
        $j = [IO.File]::ReadAllText($picks) | ConvertFrom-Json
        Write-Log "  today_picks.json: asOf=$($j.asOf) buyDate=$($j.buyDate) generatedAt=$($j.generatedAt)"
        if ($j.asOf -ne $targetStr) { $problems.Add("today_picks.json の asOf が $($j.asOf) (期待 $targetStr)") }
    }
    return $problems
}

Write-Log "=== 取得の確認 開始 (対象 $targetStr $(Format-JpDate -Date $target) / 実行 $($now.ToString('yyyy-MM-dd HH:mm'))) ==="

# 対象日は必ず営業日になるが、-AsOf で休業日を渡されたときだけ知らせる
if (-not (Test-TradingDay -Date $target)) {
    Write-Log "$targetStr は東証の休業日。-AsOf の指定を確認すること" "WARN"
}

# 17時の実行が何を報告しているか
$runStatus = Join-Path $logDirFull "last_run.json"
if (Test-Path $runStatus) {
    $r = [IO.File]::ReadAllText($runStatus) | ConvertFrom-Json
    Write-Log "  last_run.json: date=$($r.date) ok=$($r.ok) error=$($r.error)"
    if ($r.date -ne $targetStr) { Write-Log "  17時の実行記録が $targetStr のものではない" "WARN" }
} else {
    Write-Log "  last_run.json がない(17時の更新が一度も動いていない)" "WARN"
}

$problems = Test-Freshness
$repaired = $false

if ($problems.Count -eq 0) {
    Write-Log "=== 正常: $targetStr のデータはそろっている ==="
} else {
    foreach ($p in $problems) { Write-Log "異常: $p" "WARN" }
    # 取引時間中(9:00〜16:00)に取ると最新行が途中の値になるので手を出さない。
    $inMarket = (Test-TradingDay -Date $now.Date) -and $now.Hour -ge 9 -and $now.Hour -lt 16
    if ($NoRepair) {
        Write-Log "NoRepair が指定されているので取り直さない"
    } elseif ($inMarket -and -not $Force) {
        Write-Log "取引時間中なので取り直さない。17時の更新を待つこと(急ぐなら -Force)" "WARN"
    } else {
        Write-Log "--- 取り直しを実行する"
        # Invoke-DailyUpdate は当日基準で動くので、前営業日の分を取りに行くため -Force を渡す
        & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root "ps\Invoke-DailyUpdate.ps1") -Budget $Budget -Force *>> $logPath
        $code = $LASTEXITCODE
        $repaired = $true
        Write-Log "取り直し終了 (exit $code)"
        $problems = Test-Freshness
        if ($problems.Count -eq 0) { Write-Log "=== 復旧した ===" }
        else { foreach ($p in $problems) { Write-Log "復旧できず: $p" "ERROR" } }
    }
}

$out = [PSCustomObject]@{
    date = $targetStr; checkedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    ok = ($problems.Count -eq 0); skipped = $false; repaired = $repaired; problems = @($problems)
}
[IO.File]::WriteAllText($statusPath, ($out | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding $false))

if ($problems.Count -eq 0) { exit 0 } else { exit 1 }
