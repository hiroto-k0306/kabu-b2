# 共通関数（設定読み込み、パス解決、Yahoo Financeからのデータ取得、数値ユーティリティ）

function Get-ProjectRoot {
    (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}

function Get-Config {
    # 環境変数 KABU_CONFIG があればその設定ファイルを使う（並列検証で設定を切り替えるため）
    $path = Join-Path $PSScriptRoot "config.json"
    if (-not [string]::IsNullOrWhiteSpace($env:KABU_CONFIG)) {
        $path = $env:KABU_CONFIG
        if (-not [System.IO.Path]::IsPathRooted($path)) { $path = Join-Path (Get-ProjectRoot) $path }
    }
    Get-Content -Path $path -Raw -Encoding UTF8 | ConvertFrom-Json
}

$script:PriceModes = @("close_to_close", "close_to_open", "open_to_close", "open_to_open", "close_to_close_lag1", "close_to_open_lag1")

function Get-PriceMode {
    param([Parameter(Mandatory)]$Config)
    $mode = [string]$Config.label.priceMode
    if ([string]::IsNullOrWhiteSpace($mode)) { $mode = "close_to_close" }
    if ($script:PriceModes -notcontains $mode) {
        throw "unknown label.priceMode '$mode' (expected one of: $($script:PriceModes -join ', '))"
    }
    return $mode
}

function Get-PriceModeDescription {
    param([Parameter(Mandatory)][string]$Mode)
    switch ($Mode) {
        "close_to_close" { "当日引けで買い、翌営業日の引けで売る" }
        "close_to_open" { "当日引けで買い、翌営業日の寄付きで売る" }
        "open_to_close" { "翌営業日の寄付きで買い、同日の引けで売る" }
        "open_to_open" { "翌営業日の寄付きで買い、翌々営業日の寄付きで売る" }
        "close_to_close_lag1" { "翌営業日の引けで買い、翌々営業日の引けで売る（前日までのデータで判断）" }
        "close_to_open_lag1" { "翌営業日の引けで買い、翌々営業日の寄付きで売る（前日までのデータで判断）" }
    }
}

function Get-SeriesStats {
    <#
        日次リターン列から累積・年率・Sharpe・平均日次リターンのt値・勝率・最大DDを計算する。
        NaNの行は無視する。t値は日々のリターンを独立とみなした簡易検定。
    #>
    param(
        [Parameter(Mandatory)][object[]]$Rows,
        [Parameter(Mandatory)][string]$Field
    )
    $cum = 1.0
    $curve = New-Object System.Collections.Generic.List[double]
    $sumRet = 0.0; $sumSq = 0.0
    $wins = 0; $n = 0
    $peak = 1.0; $maxDD = 0.0
    foreach ($r in $Rows) {
        $v = [double]$r.$Field
        if ([double]::IsNaN($v)) { continue }
        $n++
        $cum = $cum * (1 + $v)
        $curve.Add($cum)
        $sumRet += $v
        $sumSq += $v * $v
        if ($v -gt 0) { $wins++ }
        if ($cum -gt $peak) { $peak = $cum }
        $dd = ($cum - $peak) / $peak
        if ($dd -lt $maxDD) { $maxDD = $dd }
    }
    $meanRet = 0.0; $sd = 0.0; $sharpe = 0.0; $t = 0.0; $winRate = 0.0; $annualized = 0.0
    if ($n -gt 0 -and $cum -gt 0) { $annualized = [Math]::Pow($cum, 252.0 / $n) - 1.0 }
    if ($n -gt 1) {
        $meanRet = $sumRet / $n
        $variance = ($sumSq - $n * $meanRet * $meanRet) / ($n - 1)
        if ($variance -gt 0) { $sd = [Math]::Sqrt($variance) }
        if ($sd -gt 0) {
            $sharpe = ($meanRet / $sd) * [Math]::Sqrt(252)
            $t = $meanRet / ($sd / [Math]::Sqrt($n))
        }
        $winRate = $wins / $n
    }
    [PSCustomObject]@{
        curve       = $curve
        days        = $n
        totalReturn = $cum - 1.0
        annualized  = $annualized
        meanDaily   = $meanRet
        sharpe      = $sharpe
        tStat       = $t
        winRate     = $winRate
        maxDD       = $maxDD
    }
}

function Import-IndexMembership {
    <#
        universe.changesCsv (date,action,code,name) があれば、各銘柄が指数に入っていた期間を返す。
        code -> {start; end}。start/end が空文字なら期間の制限なし。start <= 日付 < end の日だけ構成銘柄。
        universe.csv は期間中に一度でも構成銘柄だった全銘柄を含むこと。changesCsv が無ければ $null。
    #>
    param([Parameter(Mandatory)]$Config)
    if (-not $Config.universe.changesCsv) { return $null }
    $m = @{}
    foreach ($u in (Import-Csv -Path (Resolve-ProjectPath $Config.universe.csv) -Encoding UTF8)) {
        $m[$u.code] = [PSCustomObject]@{ start = ""; end = "" }
    }
    foreach ($ch in (Import-Csv -Path (Resolve-ProjectPath $Config.universe.changesCsv) -Encoding UTF8)) {
        if (-not $m.ContainsKey($ch.code)) { throw "changesCsv の $($ch.code) が universe.csv にありません" }
        $e = $m[$ch.code]
        if ($ch.action -eq "add") {
            if ($e.start) { throw "$($ch.code) に採用が複数あります（1銘柄1期間のみ対応）" }
            $e.start = $ch.date
        } elseif ($ch.action -eq "remove") {
            if ($e.end) { throw "$($ch.code) に除外が複数あります（1銘柄1期間のみ対応）" }
            $e.end = $ch.date
        } else { throw "unknown action: $($ch.action)" }
        if ($e.start -and $e.end -and [string]::Compare($e.end, $e.start) -le 0) { throw "$($ch.code) は除外後の再採用で、1期間では表せません" }
    }
    return $m
}

function Test-IndexMember {
    param($Membership, [string]$Code, [string]$Date)
    if ($null -eq $Membership) { return $true }
    if (-not $Membership.ContainsKey($Code)) { return $false }
    $e = $Membership[$Code]
    if ($e.start -and [string]::Compare($Date, $e.start) -lt 0) { return $false }
    if ($e.end -and [string]::Compare($Date, $e.end) -ge 0) { return $false }
    return $true
}

function Import-TopRanking {
    <#
        日ごとの売買代金ランキング(data.dailyRankingCsv)から上位 universe.topNPopular 件を読み、
        date -> 銘柄コードのList を返す。skabu.startDate があればそれ以降の日だけ。
    #>
    param([Parameter(Mandatory)]$Config)
    $topN = [int]$Config.universe.topNPopular
    $startDate = ""
    if ($Config.skabu -and $Config.skabu.startDate) { $startDate = [string]$Config.skabu.startDate }
    $ranking = @{}
    foreach ($row in (Import-Csv -Path (Resolve-ProjectPath $Config.data.dailyRankingCsv) -Encoding UTF8)) {
        if ([int]$row.rank -gt $topN) { continue }
        if ($startDate -and [string]::Compare($row.date, $startDate) -lt 0) { continue }
        if (-not $ranking.ContainsKey($row.date)) { $ranking[$row.date] = New-Object System.Collections.Generic.List[string] }
        $ranking[$row.date].Add($row.code)
    }
    return $ranking
}

function Import-PriceBook {
    <#
        指定した銘柄だけ、調整済みの始値・終値と、分割・配当調整前の実際の始値・高値・安値・終値を読み込む。
        skabu.tickerDir があれば銘柄別ファイル（Fetch-UniversePrices.ps1）から、
        無ければ data.pricesCsv と skabu.rawPricesCsv から読む。
        戻り値: px (code -> 系列) と calendar (全銘柄の日付の和集合、昇順)
    #>
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string[]]$Codes)
    $codeSet = New-Object System.Collections.Generic.HashSet[string]
    foreach ($c in $Codes) { [void]$codeSet.Add($c) }
    $px = @{}
    $dateSet = New-Object System.Collections.Generic.HashSet[string]

    $newSeries = {
        [PSCustomObject]@{
            idx      = New-Object 'System.Collections.Generic.Dictionary[string,int]'
            adjOpen  = New-Object System.Collections.Generic.List[double]
            adjClose = New-Object System.Collections.Generic.List[double]
            rawOpen  = New-Object System.Collections.Generic.List[double]
            rawHigh  = New-Object System.Collections.Generic.List[double]
            rawLow   = New-Object System.Collections.Generic.List[double]
            rawClose = New-Object System.Collections.Generic.List[double]
        }
    }

    if ($Config.skabu.tickerDir) {
        $dir = Split-Path (Resolve-ProjectPath "$($Config.skabu.tickerDir)/placeholder") -Parent
        # skabu.extraTickerDir: tickerDir に無い銘柄（インバースETFなど）を探す場所
        $extraDir = $null
        if ($Config.skabu.extraTickerDir) { $extraDir = Split-Path (Resolve-ProjectPath "$($Config.skabu.extraTickerDir)/placeholder") -Parent }
        # skabu.priceFromDate: これより前の行は読まない（長期データの読み込みを速くするため）
        $priceFrom = ""
        if ($Config.skabu.priceFromDate) { $priceFrom = [string]$Config.skabu.priceFromDate }
        foreach ($code in $codeSet) {
            $path = Join-Path $dir "$code.csv"
            if (-not (Test-Path $path) -and $extraDir) { $path = Join-Path $extraDir "$code.csv" }
            if (-not (Test-Path $path)) { Write-Warning "price file missing: $path"; continue }
            $p = & $newSeries
            $px[$code] = $p
            foreach ($row in (Import-Csv -Path $path -Encoding UTF8)) {
                if ($priceFrom -and [string]::CompareOrdinal($row.date, $priceFrom) -lt 0) { continue }
                $p.idx[$row.date] = $p.adjOpen.Count
                $p.adjOpen.Add([double]$row.open)
                $p.adjClose.Add([double]$row.close)
                $p.rawOpen.Add([double]$row.raw_open)
                $p.rawHigh.Add([double]$row.raw_high)
                $p.rawLow.Add([double]$row.raw_low)
                $p.rawClose.Add([double]$row.raw_close)
                [void]$dateSet.Add($row.date)
            }
        }
    } else {
        foreach ($row in (Import-Csv -Path (Resolve-ProjectPath $Config.data.pricesCsv) -Encoding UTF8)) {
            if (-not $codeSet.Contains($row.code)) { continue }
            if (-not $px.ContainsKey($row.code)) { $px[$row.code] = & $newSeries }
            $p = $px[$row.code]
            $p.idx[$row.date] = $p.adjOpen.Count
            $p.adjOpen.Add([double]$row.open)
            $p.adjClose.Add([double]$row.close)
            $p.rawOpen.Add([double]::NaN)
            $p.rawHigh.Add([double]::NaN)
            $p.rawLow.Add([double]::NaN)
            $p.rawClose.Add([double]::NaN)
            [void]$dateSet.Add($row.date)
        }
        foreach ($row in (Import-Csv -Path (Resolve-ProjectPath $Config.skabu.rawPricesCsv) -Encoding UTF8)) {
            if (-not $px.ContainsKey($row.code)) { continue }
            $p = $px[$row.code]
            $i = 0
            if ($p.idx.TryGetValue($row.date, [ref]$i)) {
                $p.rawOpen[$i] = [double]$row.raw_open
                $p.rawHigh[$i] = [double]$row.raw_high
                $p.rawLow[$i] = [double]$row.raw_low
                $p.rawClose[$i] = [double]$row.raw_close
            }
        }
    }

    foreach ($code in @($px.Keys)) {
        $n = Set-InvalidPriceRows -Series $px[$code]
        if ($n -gt 0) { Write-Warning "$code : $n rows marked invalid (corrupted values or discontinuity)" }
    }

    [PSCustomObject]@{ px = $px; calendar = @($dateSet | Sort-Object) }
}

$script:MaxValidPrice = 5000000.0
$script:MinOvernightRatio = 0.4
$script:MaxOvernightRatio = 2.5
$script:MaxDateGapDays = 30

function Test-PriceRowValid {
    param([double]$Open, [double]$Close, [double]$RawClose, [double]$Volume = 1.0)
    if (-not ($Open -gt 0) -or -not ($Close -gt 0)) { return $false }
    if ($Open -gt $script:MaxValidPrice -or $Close -gt $script:MaxValidPrice) { return $false }
    if (-not [double]::IsNaN($RawClose) -and ($RawClose -lt 1 -or $RawClose -gt $script:MaxValidPrice)) { return $false }
    if (-not ($Volume -gt 0)) { return $false }
    return $true
}

function Set-InvalidPriceRows {
    <#
        Yahooのデータには、上場廃止後に同じコードで再上場した銘柄が1本につながって値が壊れていたり、
        株式分割の調整漏れで不連続になっていたりするものがある。
        値が不正な行と、前の行から不連続になっている行を NaN にして、騰落率の計算に使われないようにする。
        （日本株は1日の値幅制限があるので、前日終値→当日始値が0.4倍未満/2.5倍超になることは通常無い）
        戻り値: NaN にした行数
    #>
    param([Parameter(Mandatory)]$Series)
    $count = $Series.adjOpen.Count
    $dates = [string[]]::new($count)
    foreach ($kv in $Series.idx.GetEnumerator()) { $dates[$kv.Value] = $kv.Key }
    $marked = 0
    $lastValid = -1
    for ($i = 0; $i -lt $count; $i++) {
        $o = $Series.adjOpen[$i]
        $valid = Test-PriceRowValid -Open $o -Close $Series.adjClose[$i] -RawClose $Series.rawClose[$i]
        $isBreak = $false
        if ($valid -and $lastValid -ge 0) {
            if (([datetime]$dates[$i] - [datetime]$dates[$lastValid]).TotalDays -gt $script:MaxDateGapDays) { $isBreak = $true }
            elseif ($lastValid -eq ($i - 1)) {
                $ratio = $o / $Series.adjClose[$lastValid]
                if ($ratio -gt $script:MaxOvernightRatio -or $ratio -lt $script:MinOvernightRatio) { $isBreak = $true }
            }
        }
        if ($valid -and -not $isBreak) { $lastValid = $i; continue }
        # 不正な行、または不連続の直後の行は使わない。不連続の場合はその次の行から新しい系列として扱う
        if ($isBreak) { $lastValid = -1 }
        $Series.adjOpen[$i] = [double]::NaN
        $Series.adjClose[$i] = [double]::NaN
        $Series.rawOpen[$i] = [double]::NaN
        $Series.rawHigh[$i] = [double]::NaN
        $Series.rawLow[$i] = [double]::NaN
        $Series.rawClose[$i] = [double]::NaN
        $marked++
    }
    return $marked
}

function Get-PowerShellExe {
    # 子スクリプトを今と同じ PowerShell で動かす(Windows の powershell.exe / Linux の pwsh)
    (Get-Process -Id $PID).Path
}

function ConvertTo-JsonNumber {
    # 整数値の double は long にして返す。PowerShell 7 の ConvertTo-Json は double の 500000 を
    # 500000.0 と書くので、Windows PowerShell 5.1 と同じ出力にするために使う
    param([double]$Value)
    if ($Value -eq [Math]::Floor($Value) -and [Math]::Abs($Value) -lt 1e15) { return [long]$Value }
    return $Value
}

function Resolve-ProjectPath {
    param([Parameter(Mandatory)][string]$RelativePath)
    $full = Join-Path (Get-ProjectRoot) $RelativePath
    $dir = Split-Path $full -Parent
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $full
}

function Get-YahooChart {
    <#
        Yahoo Financeの公開chart APIから日足OHLCVを取得する。
        yfinanceライブラリが内部で使っているものと同じエンドポイント。
        auto_adjust相当: adjclose/close の比率でOHLCを調整する。
    #>
    param(
        [Parameter(Mandatory)][string]$Code,
        [int]$Years = 2,
        # 指定すると、分割も配当も調整していない当時の実際の価格(raw_open/raw_high/raw_low/raw_close)も返す。
        # APIのopen/closeは分割調整済みなので、分割日より前の足に分割比率を掛けて戻す。
        [switch]$WithRawPrices,
        # 指定すると Years ではなくこの日付(yyyy-MM-dd)から現在までを取得する（range=max は月足になるため period1/period2 を使う）
        [string]$StartDate = ""
    )
    $symbol = "$Code.T"
    $uri = "https://query1.finance.yahoo.com/v8/finance/chart/${symbol}?range=${Years}y&interval=1d"
    if ($StartDate) {
        $p1 = ([DateTimeOffset][datetime]::SpecifyKind([datetime]$StartDate, "Utc")).ToUnixTimeSeconds()
        $p2 = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
        $uri = "https://query1.finance.yahoo.com/v8/finance/chart/${symbol}?period1=$p1&period2=$p2&interval=1d"
    }
    if ($WithRawPrices) { $uri += "&events=split" }
    $headers = @{ "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" }

    $resp = $null
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            $resp = Invoke-RestMethod -Uri $uri -Headers $headers -TimeoutSec 20
            break
        } catch {
            Start-Sleep -Seconds (2 * $attempt)
        }
    }
    if ($null -eq $resp) {
        Write-Warning "  fetch failed for $symbol (all retries exhausted)"
        return @()
    }

    $result = $resp.chart.result
    if (-not $result) { return @() }
    $r = $result[0]
    $ts = $r.timestamp
    if (-not $ts) { return @() }

    $quote = $r.indicators.quote[0]
    $adjArr = $null
    if ($r.indicators.adjclose) { $adjArr = $r.indicators.adjclose[0].adjclose }

    $splits = @()
    if ($WithRawPrices -and $r.events -and $r.events.splits) {
        $splits = @($r.events.splits.PSObject.Properties | ForEach-Object {
            $s = $_.Value
            [PSCustomObject]@{
                date  = [DateTimeOffset]::FromUnixTimeSeconds([int64]$s.date).UtcDateTime.ToString("yyyy-MM-dd")
                ratio = [double]$s.numerator / [double]$s.denominator
            }
        })
    }

    $rows = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $ts.Count; $i++) {
        $c = $quote.close[$i]
        $o = $quote.open[$i]
        $h = $quote.high[$i]
        $l = $quote.low[$i]
        $v = $quote.volume[$i]
        if ($null -eq $c -or $null -eq $o -or $null -eq $h -or $null -eq $l -or $null -eq $v) { continue }
        $adjClose = if ($adjArr -and $null -ne $adjArr[$i]) { [double]$adjArr[$i] } else { [double]$c }
        $ratio = if ([double]$c -ne 0) { $adjClose / [double]$c } else { 1.0 }
        $date = [DateTimeOffset]::FromUnixTimeSeconds([int64]$ts[$i]).UtcDateTime.Date

        $dateStr = $date.ToString("yyyy-MM-dd")
        $row = [ordered]@{
            date   = $dateStr
            code   = $Code
            open   = [double]$o * $ratio
            high   = [double]$h * $ratio
            low    = [double]$l * $ratio
            close  = $adjClose
            volume = [double]$v
        }
        if ($WithRawPrices) {
            # 分割日(権利落ち日)より前の足は、その分割比率だけ実際の価格が高かった
            $splitFactor = 1.0
            foreach ($s in $splits) { if ([string]::Compare($dateStr, $s.date) -lt 0) { $splitFactor *= $s.ratio } }
            $row["raw_open"] = [double]$o * $splitFactor
            $row["raw_high"] = [double]$h * $splitFactor
            $row["raw_low"] = [double]$l * $splitFactor
            $row["raw_close"] = [double]$c * $splitFactor
        }
        # 極端な分割や異常値の銘柄では調整比率が破綻し、±∞ になることがある。
        # そのまま書くと PowerShell 側が [double] に読み戻せず落ちるので、NaN にして
        # Test-PriceRowValid に「不正な行」として捨てさせる(C#側の ParseD も NaN 扱い)。
        foreach ($key in @($row.Keys)) {
            if ($key -eq "date" -or $key -eq "code") { continue }
            $val = [double]$row[$key]
            if ([double]::IsInfinity($val)) { $row[$key] = [double]::NaN }
        }
        $rows.Add([PSCustomObject]$row)
    }
    return $rows
}

function Invert-Matrix {
    <#
        Gauss-Jordan法によるn×n正方行列の逆行列（小さいサイズ専用、特徴量数+1程度を想定）
        注意: PowerShellの多次元配列は $a[$i,$j] をメソッド呼び出しの引数に直接ネストすると
        パーサがカンマを引数区切りと誤認するため、必ず一度スカラー変数に取り出してから使う。
    #>
    param([double[,]]$M)
    $n = $M.GetLength(0)
    $a = [double[,]]::new($n, (2 * $n))
    for ($i = 0; $i -lt $n; $i++) {
        for ($j = 0; $j -lt $n; $j++) {
            $v = $M[$i, $j]
            $a[$i, $j] = $v
        }
        $identCol = $n + $i
        $a[$i, $identCol] = 1.0
    }
    for ($col = 0; $col -lt $n; $col++) {
        $pivot = $col
        $pivotCell = $a[$col, $col]
        $maxVal = [Math]::Abs($pivotCell)
        for ($r = $col + 1; $r -lt $n; $r++) {
            $cell = $a[$r, $col]
            $absCell = [Math]::Abs($cell)
            if ($absCell -gt $maxVal) { $maxVal = $absCell; $pivot = $r }
        }
        if ($maxVal -lt 1e-12) { continue }
        if ($pivot -ne $col) {
            for ($j = 0; $j -lt 2 * $n; $j++) {
                $tmp = $a[$col, $j]
                $swap = $a[$pivot, $j]
                $a[$col, $j] = $swap
                $a[$pivot, $j] = $tmp
            }
        }
        $pivotVal = $a[$col, $col]
        for ($j = 0; $j -lt 2 * $n; $j++) {
            $cur = $a[$col, $j]
            $a[$col, $j] = $cur / $pivotVal
        }
        for ($r = 0; $r -lt $n; $r++) {
            if ($r -eq $col) { continue }
            $factor = $a[$r, $col]
            if ($factor -eq 0) { continue }
            for ($j = 0; $j -lt 2 * $n; $j++) {
                $curR = $a[$r, $j]
                $curCol = $a[$col, $j]
                $a[$r, $j] = $curR - $factor * $curCol
            }
        }
    }
    $inv = [double[,]]::new($n, $n)
    for ($i = 0; $i -lt $n; $i++) {
        for ($j = 0; $j -lt $n; $j++) {
            $srcCol = $n + $j
            $v2 = $a[$i, $srcCol]
            $inv[$i, $j] = $v2
        }
    }
    return , $inv
}


# --- 東証の営業日 ---
# 休業日 = 土日 + 祝日 + 年末年始(12/31〜1/3)。祝日は法律の規則から作った内蔵リストなので、
# 年をまたぐ前に取引所のカレンダーで確認して足すこと。
$script:JpxHolidays = @(
    "2026-01-01", "2026-01-02", "2026-01-12", "2026-02-11", "2026-02-23", "2026-03-20",
    "2026-04-29", "2026-05-03", "2026-05-04", "2026-05-05", "2026-05-06", "2026-07-20",
    "2026-08-11", "2026-09-21", "2026-09-22", "2026-09-23", "2026-10-12", "2026-11-03",
    "2026-11-23", "2026-12-31",
    "2027-01-01", "2027-01-02", "2027-01-03", "2027-01-11", "2027-02-11", "2027-02-23",
    "2027-03-21", "2027-03-22", "2027-04-29", "2027-05-03", "2027-05-04", "2027-05-05",
    "2027-07-19", "2027-08-11", "2027-09-20", "2027-09-23", "2027-10-11", "2027-11-03",
    "2027-11-23", "2027-12-31"
)
$script:JpxHolidayMaxDate = ($script:JpxHolidays | Sort-Object)[-1]

# 曜日の日本語表記。dot-source した側でもそのまま $dowJa として使える。
$dowJa = @{ "Sunday" = "日"; "Monday" = "月"; "Tuesday" = "火"; "Wednesday" = "水"; "Thursday" = "木"; "Friday" = "金"; "Saturday" = "土" }

function Format-JpDate {
    <# 2026-09-24 -> 9月24日(木) #>
    param([Parameter(Mandatory)][datetime]$Date)
    "{0}月{1}日({2})" -f $Date.Month, $Date.Day, $dowJa[$Date.DayOfWeek.ToString()]
}

function Get-JpxHolidays { return $script:JpxHolidays }

function Test-TradingDay {
    <# 東証が開いている日なら $true。内蔵の祝日リストを超える日付は警告を出す。 #>
    param([Parameter(Mandatory)][datetime]$Date)
    if ($Date.DayOfWeek -eq "Saturday" -or $Date.DayOfWeek -eq "Sunday") { return $false }
    $s = $Date.ToString("yyyy-MM-dd")
    if ([string]::Compare($s, $script:JpxHolidayMaxDate) -gt 0) {
        Write-Warning "$s は内蔵の祝日リスト(〜$script:JpxHolidayMaxDate)の範囲外。Common.ps1 の JpxHolidays に追記すること"
    }
    return -not ($script:JpxHolidays -contains $s)
}

function Get-NextTradingDay {
    <# Date の翌営業日を返す #>
    param([Parameter(Mandatory)][datetime]$Date)
    for ($i = 1; $i -le 30; $i++) {
        $x = $Date.AddDays($i)
        if (Test-TradingDay -Date $x) { return $x }
    }
    return $Date.AddDays(1)
}

function Get-PrevTradingDay {
    <# Date の前営業日を返す #>
    param([Parameter(Mandatory)][datetime]$Date)
    for ($i = 1; $i -le 30; $i++) {
        $x = $Date.AddDays(-$i)
        if (Test-TradingDay -Date $x) { return $x }
    }
    return $Date.AddDays(-1)
}
