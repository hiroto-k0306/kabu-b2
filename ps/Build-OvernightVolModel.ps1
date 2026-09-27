# B2が実際に負う「引け→翌寄りの窓」の大きさを予測するモデル。
#
# 別プロジェクト(kabu)の変動率モデルは日経平均の日中変動(高安値)を予測する。B2が負うのは
# 買った日の引けから翌営業日の寄りまでの窓なので、対象が違う。ここではその窓を直接予測する。
#
# 予測の対象: 選定日 t の上位10銘柄を、翌営業日 D の引けで等分に買い、その翌営業日 S の寄りで
#            売ったときの騰落率 r の2乗 (%^2)。r = 10銘柄の (寄り/前引け - 1) の平均。
# 使う情報  : t の引けまでに分かっているものだけ。加えて D と S は取引所カレンダーから
#            事前に分かるので、窓の暦日数(連休をまたぐか)は説明変数に使える。
#
# 出力:
#   reports/overnight_vol/dataset.csv      目的変数と説明変数
#   reports/overnight_vol/predictions.csv  volSizing がそのまま読める形
#   reports/overnight_vol/model_<年>.json  年ごとの係数
#
# 使い方: $env:KABU_CONFIG = "ps/config.overnight_vol.json"; powershell -File ps\Build-OvernightVolModel.ps1
param(
    [int]$FirstPredictYear = 2024,   # この年から予測する(それ以前は学習に使う)
    [int]$MinTrain = 300,            # 学習に必要な最小日数
    [double]$Floor = 0.001,          # 目的変数の下限(%^2)。対数を取るため
    [double]$EwmaLambda = 0.94,      # 比較用EWMAの減衰
    [int]$BasketWindow = 20,         # 銘柄ごとの窓の大きさを測る日数
    [int]$MarketWindow = 22,         # 日経平均の変動率を測る日数
    # 水準の測り直しは「予測/実現」を1に近づけるが、裾の重い目的変数ではQLIKEが悪化する。
    # サイジングに使うだけなら倍率は規格化されるので、-NoRecal の方が予測精度は良い。
    [switch]$NoRecal,
    [int]$RecalWindow = 250,         # 水準を測り直す窓(直近何日の実績で補正し直すか)
    [int]$RecalMinObs = 60,          # 測り直しに必要な最小日数
    [double]$RecalMin = 0.5,         # 測り直しの倍率の下限
    [double]$RecalMax = 2.0,         # 同 上限
    [string]$OutDir = "reports/overnight_vol"
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"
Set-Location (Get-ProjectRoot)

$config = Get-Config
$ranking = Import-TopRanking -Config $config
if ($ranking.Count -eq 0) { throw "選定結果が読めない" }

$codeSet = New-Object System.Collections.Generic.HashSet[string]
foreach ($k in $ranking.Keys) { foreach ($c in $ranking[$k]) { [void]$codeSet.Add($c) } }
$codes = @($codeSet)
Write-Host "選定日 $($ranking.Count)日 / 銘柄 $($codes.Count)"

$book = Import-PriceBook -Config $config -Codes $codes
$px = $book.px
$calendar = $book.calendar
foreach ($c in $codes) { if ($px.ContainsKey($c)) { [void](Set-InvalidPriceRows -Series $px[$c]) } }
Write-Host "カレンダー $($calendar[0]) - $($calendar[-1]) ($($calendar.Count)日)"

# カレンダー上の位置
$calIdx = New-Object 'System.Collections.Generic.Dictionary[string,int]'
for ($i = 0; $i -lt $calendar.Count; $i++) { $calIdx[[string]$calendar[$i]] = $i }

# --- 銘柄ごとに「窓の騰落率」をカレンダーに合わせて並べ、累積和を作る(窓の統計をO(1)で出す) ---
# on[i] = 銘柄の calendar[i] の寄り ÷ calendar[i-1] の引け - 1
$onRet = @{}
$onCnt = @{}
$onSum = @{}
$onSq  = @{}
foreach ($c in $codes) {
    if (-not $px.ContainsKey($c)) { continue }
    $ser = $px[$c]
    $n = $calendar.Count
    $r = [double[]]::new($n)
    for ($i = 0; $i -lt $n; $i++) { $r[$i] = [double]::NaN }
    for ($i = 1; $i -lt $n; $i++) {
        $d0 = [string]$calendar[$i - 1]; $d1 = [string]$calendar[$i]
        if (-not $ser.idx.ContainsKey($d0) -or -not $ser.idx.ContainsKey($d1)) { continue }
        $j0 = $ser.idx[$d0]; $j1 = $ser.idx[$d1]
        $pc = $ser.adjClose[$j0]; $po = $ser.adjOpen[$j1]
        if ([double]::IsNaN($pc) -or [double]::IsNaN($po) -or $pc -le 0 -or $po -le 0) { continue }
        $r[$i] = $po / $pc - 1.0
    }
    $cnt = [int[]]::new($n + 1); $sum = [double[]]::new($n + 1); $sq = [double[]]::new($n + 1)
    for ($i = 0; $i -lt $n; $i++) {
        $v = $r[$i]
        $ok = -not [double]::IsNaN($v)
        $cnt[$i + 1] = $cnt[$i] + $(if ($ok) { 1 } else { 0 })
        $sum[$i + 1] = $sum[$i] + $(if ($ok) { $v } else { 0.0 })
        $sq[$i + 1]  = $sq[$i]  + $(if ($ok) { $v * $v } else { 0.0 })
    }
    $onRet[$c] = $r; $onCnt[$c] = $cnt; $onSum[$c] = $sum; $onSq[$c] = $sq
}

function Get-WindowRms {
    # calendar の [$from..$to] における窓騰落率の二乗平均平方根(%)。件数が足りなければ NaN
    param([string]$Code, [int]$From, [int]$To, [int]$MinN = 5)
    if ($From -lt 0) { $From = 0 }
    if ($To -lt $From) { return [double]::NaN }
    $cnt = $onCnt[$Code]; $sq = $onSq[$Code]
    $n = $cnt[$To + 1] - $cnt[$From]
    if ($n -lt $MinN) { return [double]::NaN }
    $ss = $sq[$To + 1] - $sq[$From]
    return 100.0 * [Math]::Sqrt($ss / $n)
}

# --- 日経平均: 終値の変動率と、寄りの窓の大きさ ---
$nk = @{}
$nkDates = New-Object System.Collections.Generic.List[string]
foreach ($row in (Import-Csv -Path (Resolve-ProjectPath "data/raw/market/daily_since2000/N225.csv") -Encoding UTF8)) {
    $o = 0.0; $c = 0.0
    if (-not [double]::TryParse([string]$row.open, [ref]$o)) { continue }
    if (-not [double]::TryParse([string]$row.close, [ref]$c)) { continue }
    if ($o -le 0 -or $c -le 0) { continue }
    $d = [string]$row.date
    $nk[$d] = [PSCustomObject]@{ open = $o; close = $c }
    $nkDates.Add($d)
}
$nkSorted = @($nkDates | Sort-Object)
$nkPos = New-Object 'System.Collections.Generic.Dictionary[string,int]'
for ($i = 0; $i -lt $nkSorted.Count; $i++) { $nkPos[[string]$nkSorted[$i]] = $i }

function Get-NikkeiStats {
    # 日付 $D までの直近 $W 日で、終値の変動率(%) と 寄りの窓の二乗平均平方根(%) を返す
    param([string]$AsOf, [int]$W)
    if (-not $nkPos.ContainsKey($AsOf)) { return @([double]::NaN, [double]::NaN) }
    $p = $nkPos[$AsOf]
    $ccSq = 0.0; $ccN = 0; $gpSq = 0.0; $gpN = 0
    for ($i = $p; $i -gt 0 -and ($p - $i) -lt $W; $i--) {
        $a = $nk[$nkSorted[$i - 1]]; $b = $nk[$nkSorted[$i]]
        $cc = $b.close / $a.close - 1.0
        $gp = $b.open / $a.close - 1.0
        $ccSq += $cc * $cc; $ccN++
        $gpSq += $gp * $gp; $gpN++
    }
    $cc = if ($ccN -ge 5) { 100.0 * [Math]::Sqrt($ccSq / $ccN) } else { [double]::NaN }
    $gp = if ($gpN -ge 5) { 100.0 * [Math]::Sqrt($gpSq / $gpN) } else { [double]::NaN }
    return @($cc, $gp)
}

# --- 1行 = 1買付日 ---
$rows = New-Object System.Collections.Generic.List[object]
foreach ($t in ($ranking.Keys | Sort-Object)) {
    $ts = [string]$t
    if (-not $calIdx.ContainsKey($ts)) { continue }
    $i = $calIdx[$ts]
    if (($i + 2) -ge $calendar.Count) { continue }
    $buyDate  = [string]$calendar[$i + 1]   # 引けで買う日
    $sellDate = [string]$calendar[$i + 2]   # 寄りで売る日

    # 目的変数: 10銘柄を等分に持ったときの窓の騰落率
    $picks = @($ranking[$ts])
    $sum = 0.0; $n = 0
    foreach ($c in $picks) {
        if (-not $px.ContainsKey($c)) { continue }
        $ser = $px[$c]
        if (-not $ser.idx.ContainsKey($buyDate) -or -not $ser.idx.ContainsKey($sellDate)) { continue }
        $pc = $ser.adjClose[$ser.idx[$buyDate]]; $po = $ser.adjOpen[$ser.idx[$sellDate]]
        if ([double]::IsNaN($pc) -or [double]::IsNaN($po) -or $pc -le 0 -or $po -le 0) { continue }
        $sum += ($po / $pc - 1.0); $n++
    }
    if ($n -lt [Math]::Ceiling($picks.Count / 2.0)) { continue }
    $r = $sum / $n
    $rv = (100.0 * $r) * (100.0 * $r)

    # 説明変数: t の引けまでの情報 + 事前に分かるカレンダー
    $bsum = 0.0; $bn = 0
    foreach ($c in $picks) {
        if (-not $onCnt.ContainsKey($c)) { continue }
        $v = Get-WindowRms -Code $c -From ($i - $BasketWindow + 1) -To $i
        if (-not [double]::IsNaN($v)) { $bsum += $v; $bn++ }
    }
    if ($bn -eq 0) { continue }
    $basket = $bsum / $bn

    $nkv = Get-NikkeiStats -AsOf $ts -W $MarketWindow
    if ([double]::IsNaN($nkv[0]) -or [double]::IsNaN($nkv[1])) { continue }

    $gapDays = ([datetime]$sellDate - [datetime]$buyDate).TotalDays
    if ($gapDays -lt 1) { $gapDays = 1 }

    $rows.Add([PSCustomObject]@{
        forecast_date = $ts
        target_date   = $buyDate
        sell_date     = $sellDate
        stocks        = $n
        overnight_ret_pct = 100.0 * $r
        rv_pct2       = $rv
        basket_rms_pct = $basket
        nk_cc_pct     = $nkv[0]
        nk_gap_pct    = $nkv[1]
        gap_days      = $gapDays
    })
}
Write-Host "使える買付日 $($rows.Count)日  $($rows[0].target_date) .. $($rows[-1].target_date)"

# --- HARラグ(過去の実現値)を足す。行は買付日の昇順 ---
$all = @($rows | Sort-Object target_date)
for ($k = 0; $k -lt $all.Count; $k++) {
    $d1 = [double]::NaN; $w = [double]::NaN; $m = [double]::NaN
    if ($k -ge 1) { $d1 = $all[$k - 1].rv_pct2 }
    if ($k -ge 5) { $s = 0.0; for ($j = $k - 5; $j -lt $k; $j++) { $s += $all[$j].rv_pct2 }; $w = $s / 5.0 }
    if ($k -ge 22) { $s = 0.0; for ($j = $k - 22; $j -lt $k; $j++) { $s += $all[$j].rv_pct2 }; $m = $s / 22.0 }
    $all[$k] | Add-Member -NotePropertyName rv_lag1 -NotePropertyValue $d1
    $all[$k] | Add-Member -NotePropertyName rv_lag5 -NotePropertyValue $w
    $all[$k] | Add-Member -NotePropertyName rv_lag22 -NotePropertyValue $m
}
$use = @($all | Where-Object { -not [double]::IsNaN($_.rv_lag22) })
Write-Host "ラグが揃う行 $($use.Count)日"

# --- 説明変数の並び ---
# 窓の暦日数は対数の連続値だと4日窓の逆転(1日窓より効率が良い)を吸収できないので区分にする。
# 基準は1〜2日。gap3 は週末をまたぐ3日、gap4p は4日以上(連休)。
$featNames = @("lrv_d", "lrv_w", "lrv_m", "lbasket", "lnk_cc", "lnk_gap", "gap3", "gap4p")
function Get-FeatureVector {
    param($Row)
    return @(
        [Math]::Log($Row.rv_lag1 + $Floor)
        [Math]::Log($Row.rv_lag5 + $Floor)
        [Math]::Log($Row.rv_lag22 + $Floor)
        [Math]::Log($Row.basket_rms_pct)
        [Math]::Log($Row.nk_cc_pct)
        [Math]::Log($Row.nk_gap_pct)
        $(if ([int]$Row.gap_days -eq 3) { 1.0 } else { 0.0 })
        $(if ([int]$Row.gap_days -ge 4) { 1.0 } else { 0.0 })
    )
}

function Invoke-Ols {
    # 標準化済みXと y から切片つきの係数を出す
    param([double[][]]$X, [double[]]$Y)
    $n = $X.Count; $p = $X[0].Count
    $m = $p + 1
    $xtx = [double[,]]::new($m, $m)
    $xty = [double[]]::new($m)
    for ($i = 0; $i -lt $n; $i++) {
        $xi = [double[]]::new($m)
        $xi[0] = 1.0
        for ($j = 0; $j -lt $p; $j++) { $xi[$j + 1] = $X[$i][$j] }
        for ($a = 0; $a -lt $m; $a++) {
            $va = $xi[$a]
            $xty[$a] += $va * $Y[$i]
            for ($b = 0; $b -lt $m; $b++) { $xtx[$a, $b] += $va * $xi[$b] }
        }
    }
    $inv = Invert-Matrix -M $xtx
    $beta = [double[]]::new($m)
    for ($a = 0; $a -lt $m; $a++) {
        $s = 0.0
        for ($b = 0; $b -lt $m; $b++) { $v = $inv[$a, $b]; $s += $v * $xty[$b] }
        $beta[$a] = $s
    }
    return $beta
}

function Get-Qlike {
    # 分散予測の誤差。小さいほうが良い
    param([double[]]$Realized, [double[]]$Pred)
    $s = 0.0; $n = 0
    for ($i = 0; $i -lt $Realized.Count; $i++) {
        $r = $Realized[$i] + $Floor; $p = $Pred[$i]
        if ($p -le 0) { continue }
        $s += $r / $p - [Math]::Log($r / $p) - 1.0
        $n++
    }
    if ($n -eq 0) { return [double]::NaN }
    return $s / $n
}

# --- 年ごとに学習して次の年を予測 ---
$outDirFull = Split-Path (Resolve-ProjectPath "$OutDir/placeholder") -Parent
New-Item -ItemType Directory -Force $outDirFull | Out-Null

$years = @($use | ForEach-Object { [int]$_.target_date.Substring(0, 4) } | Sort-Object -Unique | Where-Object { $_ -ge $FirstPredictYear })
$preds = New-Object System.Collections.Generic.List[object]

foreach ($y in $years) {
    $train = @($use | Where-Object { [int]$_.target_date.Substring(0, 4) -lt $y })
    $test  = @($use | Where-Object { [int]$_.target_date.Substring(0, 4) -eq $y })
    if ($train.Count -lt $MinTrain -or $test.Count -eq 0) { Write-Warning "$y : 学習 $($train.Count)日で足りない。飛ばす"; continue }

    $xt = @(); $yt = @()
    foreach ($r in $train) { $xt += , (Get-FeatureVector -Row $r); $yt += [Math]::Log($r.rv_pct2 + $Floor) }
    $p = $featNames.Count
    $mean = [double[]]::new($p); $sd = [double[]]::new($p)
    for ($j = 0; $j -lt $p; $j++) {
        $s = 0.0; foreach ($v in $xt) { $s += $v[$j] }
        $mean[$j] = $s / $xt.Count
        $s2 = 0.0; foreach ($v in $xt) { $d = $v[$j] - $mean[$j]; $s2 += $d * $d }
        $sd[$j] = [Math]::Sqrt($s2 / $xt.Count)
        if ($sd[$j] -le 0) { $sd[$j] = 1.0 }
    }
    $xs = @()
    foreach ($v in $xt) {
        $z = [double[]]::new($p)
        for ($j = 0; $j -lt $p; $j++) { $z[$j] = ($v[$j] - $mean[$j]) / $sd[$j] }
        $xs += , $z
    }
    $beta = Invoke-Ols -X $xs -Y ([double[]]$yt)

    # 水準補正: 学習期間で 実現平均 ÷ 予測平均 を合わせる(対数からの戻しで生じる偏りを直す)
    $sumR = 0.0; $sumP = 0.0
    for ($i = 0; $i -lt $xs.Count; $i++) {
        $f = $beta[0]
        for ($j = 0; $j -lt $p; $j++) { $f += $beta[$j + 1] * $xs[$i][$j] }
        $sumP += [Math]::Exp($f); $sumR += $train[$i].rv_pct2
    }
    $corr = if ($sumP -gt 0) { $sumR / $sumP } else { 1.0 }

    foreach ($r in $test) {
        $v = Get-FeatureVector -Row $r
        $f = $beta[0]
        for ($j = 0; $j -lt $p; $j++) { $f += $beta[$j + 1] * (($v[$j] - $mean[$j]) / $sd[$j]) }
        $pv = [Math]::Exp($f) * $corr
        if ($pv -lt $Floor) { $pv = $Floor }
        $preds.Add([PSCustomObject]@{
            forecast_date = $r.forecast_date
            target_date   = $r.target_date
            sell_date     = $r.sell_date
            model_year    = $y
            pred_raw_pct2      = $pv
            pred_variance_pct2 = $pv
            pred_sigma_pct     = [Math]::Sqrt($pv)
            realized_variance_pct2 = $r.rv_pct2
            realized_ret_pct   = $r.overnight_ret_pct
            gap_days      = $r.gap_days
            source        = "overnight_har"
        })
    }

    $model = [ordered]@{
        name = "overnight_har"; year = $y
        target = "next_open_gap_variance_of_b2_top10_pct_squared"
        features = $featNames
        mean = $mean; std = $sd
        intercept = $beta[0]
        coefficients = @($beta[1..$p])
        correction = $corr
        floor_pct_squared = $Floor
        n_train = $train.Count
        train_target_max = $train[-1].target_date
    }
    [IO.File]::WriteAllText((Join-Path $outDirFull "model_$y.json"), ([PSCustomObject]$model | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding $false))
    Write-Host ("  $y : 学習 $($train.Count)日 → 予測 $($test.Count)日  水準補正 {0:N3}" -f $corr)
}

if ($preds.Count -eq 0) { throw "予測が1件も作れなかった" }

# --- 水準の測り直し ---
# 年に一度しか係数を直さないので、変動率の水準が動く局面で予測が遅れる(2025年は予測/実現 0.61)。
# 直近 RecalWindow 日の「実現の合計 ÷ 予測の合計」を掛け直す。使うのはその日より前だけ。
$predOrder = @($preds | Sort-Object target_date)
$histR = New-Object System.Collections.Generic.List[double]
$histP = New-Object System.Collections.Generic.List[double]
foreach ($r in $predOrder) {
    $adj = 1.0
    if (-not $NoRecal -and $histR.Count -ge $RecalMinObs) {
        $from = 0
        if ($histR.Count -gt $RecalWindow) { $from = $histR.Count - $RecalWindow }
        $sr = 0.0; $sp = 0.0
        for ($i = $from; $i -lt $histR.Count; $i++) { $sr += $histR[$i]; $sp += $histP[$i] }
        if ($sp -gt 0) {
            $adj = $sr / $sp
            if ($adj -lt $RecalMin) { $adj = $RecalMin }
            if ($adj -gt $RecalMax) { $adj = $RecalMax }
        }
    }
    $pv = $r.pred_raw_pct2 * $adj
    if ($pv -lt $Floor) { $pv = $Floor }
    $r.pred_variance_pct2 = $pv
    $r.pred_sigma_pct = [Math]::Sqrt($pv)
    $r | Add-Member -NotePropertyName recal_adj -NotePropertyValue $adj -Force
    # 実績が分かってから履歴に入れる(この行の予測には使っていない)
    $histR.Add($r.realized_variance_pct2)
    $histP.Add($r.pred_raw_pct2)
}

# --- 比較用EWMA と 一定値 ---
$predSorted = @($preds | Sort-Object target_date)
$rvByDate = New-Object 'System.Collections.Generic.Dictionary[string,double]'
foreach ($r in $use) { $rvByDate[[string]$r.target_date] = $r.rv_pct2 }
$ewmaByDate = New-Object 'System.Collections.Generic.Dictionary[string,double]'
$e = $null
foreach ($r in $use) {
    if ($null -eq $e) { $e = $r.rv_pct2 } else { $ewmaByDate[[string]$r.target_date] = $e; $e = $EwmaLambda * $e + (1.0 - $EwmaLambda) * $r.rv_pct2 }
}
$firstYear = [int]$predSorted[0].target_date.Substring(0, 4)
$pre = @($use | Where-Object { [int]$_.target_date.Substring(0, 4) -lt $firstYear })
$constPred = ($pre | Measure-Object -Property rv_pct2 -Average).Average

$real = @(); $pm = @(); $pe = @(); $pc = @()
foreach ($r in $predSorted) {
    $d = [string]$r.target_date
    if (-not $ewmaByDate.ContainsKey($d)) { continue }
    $real += $r.realized_variance_pct2
    $pm += $r.pred_variance_pct2
    $pe += $ewmaByDate[$d]
    $pc += $constPred
}
$qm = Get-Qlike -Realized $real -Pred $pm
$qe = Get-Qlike -Realized $real -Pred $pe
$qc = Get-Qlike -Realized $real -Pred $pc
$biasM = (($pm | Measure-Object -Average).Average) / (($real | Measure-Object -Average).Average)
$biasE = (($pe | Measure-Object -Average).Average) / (($real | Measure-Object -Average).Average)

Write-Host ""
Write-Host "=== 窓の大きさの予測 ($($predSorted[0].target_date) 〜 $($predSorted[-1].target_date), $($real.Count)日) ==="
Write-Host ("  {0,-26} {1,10} {2,10}" -f "予測", "QLIKE", "予測/実現")
Write-Host ("  {0,-26} {1,10:N6} {2,10:N4}" -f "一定値(学習期間の平均)", $qc, ($constPred / (($real | Measure-Object -Average).Average)))
Write-Host ("  {0,-26} {1,10:N6} {2,10:N4}" -f "EWMA(λ=$EwmaLambda)", $qe, $biasE)
Write-Host ("  {0,-26} {1,10:N6} {2,10:N4}" -f "HAR+銘柄+連休(今回)", $qm, $biasM)

# EWMAとの差の有意性(Newey-Westなしの単純なt値)
$dif = @(); for ($i = 0; $i -lt $real.Count; $i++) {
    $a = $real[$i] + $Floor
    $lm = $a / $pm[$i] - [Math]::Log($a / $pm[$i]) - 1.0
    $le = $a / $pe[$i] - [Math]::Log($a / $pe[$i]) - 1.0
    $dif += ($lm - $le)
}
$dm = ($dif | Measure-Object -Average).Average
$ds = 0.0; foreach ($v in $dif) { $d0 = $v - $dm; $ds += $d0 * $d0 }
$ds = [Math]::Sqrt($ds / ($dif.Count - 1))
$tstat = $dm / ($ds / [Math]::Sqrt($dif.Count))
Write-Host ("  EWMAとの差 {0:N6}  t値 {1:N3}  (負で大きいほど改善)" -f $dm, $tstat)

# 連休の効き方
Write-Host ""
Write-Host "  窓の暦日数ごとの実現値(%^2)と予測値"
foreach ($g in ($predSorted | Group-Object gap_days | Sort-Object { [int]$_.Name })) {
    $rr = ($g.Group | Measure-Object -Property realized_variance_pct2 -Average).Average
    $pp = ($g.Group | Measure-Object -Property pred_variance_pct2 -Average).Average
    Write-Host ("    {0}日  {1,4}件  実現 {2,8:N4}  予測 {3,8:N4}" -f $g.Name, $g.Count, $rr, $pp)
}

$all | Export-Csv -Path (Join-Path $outDirFull "dataset.csv") -NoTypeInformation -Encoding UTF8
$predSorted | Export-Csv -Path (Join-Path $outDirFull "predictions.csv") -NoTypeInformation -Encoding UTF8
Write-Host ""
Write-Host "saved $OutDir/dataset.csv, predictions.csv, model_*.json"
