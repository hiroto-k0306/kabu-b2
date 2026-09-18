# SBI証券のS株(単元未満株)で「売買代金上位N銘柄を全部持つ」を実際の現金・株数で1日ずつ再現する。
#
# 再現しているルール
#   - 1株単位。1日の買付は dailyBudget まで、かつ買付余力(口座の現金)の範囲内
#   - 成行の買いは「株数×制限値幅の上限価格」で余力を拘束する。予算内に収まるよう株数を決める
#   - 株数は、上位N銘柄に予算を等分して切り捨て、余った予算で「保有額が一番少ない銘柄」に1株ずつ足す
#     （1株が予算の等分額より高い銘柄は、余りで買えなければ0株になる）
#   - close_to_open_lag1 : 前日までの確定データで選び、当日10:30〜14:00に注文 → 15:30終値で約定、
#                          翌営業日9:00始値で売却。朝の売却代金は同日の買いに使える
#   - open_to_open       : 引け後に注文 → 翌営業日9:00始値で約定、翌々営業日9:00始値で売却。
#                          売りと次の買いが同じ9:00に約定するため、売却代金は次の買いの注文時点では使えない
#   - 大引け(または寄付き)が値幅制限いっぱいで張り付いている場合は約定しないとみなす
#     （S株はストップ配分の対象外。日足からの近似）。売れなかった保有は翌営業日に再度売る
#   - 損益は「実際の株価×株数」で計算し、配当は調整済み価格の騰落率を通じて受け取ったものとみなす
#   - シナリオに "selection": { "lookback": 20, "topK": 5 } があれば、判断日までの直近 lookback 日の夜間勝率が
#     高い topK 銘柄だけに予算を等分する（ps/Test-WinRateRule.ps1 と同じルール）
#   - 税金は年ごとの損益合計がプラスなら税率を掛けて年末に差し引く（損失の繰越は考えない）
# 使い方: $env:KABU_CONFIG = "ps/config.sbi.json"; powershell -File ps\Simulate-SKabu.ps1

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"

$config = Get-Config
$sk = $config.skabu
$topN = [int]$config.universe.topNPopular
$taxRate = [double]$sk.taxRate
$outDir = Split-Path (Resolve-ProjectPath "$($sk.outputDir)/placeholder") -Parent

$limitThresholds = [double[]]@(100, 200, 500, 700, 1000, 1500, 2000, 3000, 5000, 7000, 10000, 15000, 20000, 30000, 50000, 70000, 100000, 150000, 200000, 300000, 500000, 700000, 1000000)
$limitWidths = [double[]]@(30, 50, 80, 100, 150, 300, 400, 500, 700, 1000, 1500, 3000, 4000, 5000, 7000, 10000, 15000, 30000, 40000, 50000, 70000, 100000, 150000)

function Get-LimitWidth {
    # 東証の制限値幅（基準値段=前日終値ごとの値幅）
    param([double]$Base)
    for ($k = 0; $k -lt $limitThresholds.Length; $k++) {
        if ($Base -lt $limitThresholds[$k]) { return $limitWidths[$k] }
    }
    return 300000.0
}

# --- データ読み込み ---
Write-Host "loading ranking ..."
$ranking = Import-TopRanking -Config $config
$neededCodes = @($ranking.Values | ForEach-Object { $_ } | Sort-Object -Unique)
# インバースETFで日経の値動きを打ち消すシナリオ（scenario.hedge = { code, multiple, betaCsv }）
# betaCsv の beta は判断日までに売却が済んだデータだけで推定した値（ps\Test-InverseHedge.ps1 の出力）
# ランキングCSVに weight 列があれば、銘柄ごとの予算比率として使う（日経上位10とB2上位3を半分ずつ等）
$rankWeight = @{}
$rankCsvPath = Resolve-ProjectPath $config.data.dailyRankingCsv
if ((Get-Content $rankCsvPath -TotalCount 1) -match "weight") {
    foreach ($row in (Import-Csv $rankCsvPath -Encoding UTF8)) {
        if ([int]$row.rank -gt $topN) { continue }
        if ($sk.startDate -and [string]::CompareOrdinal($row.date, [string]$sk.startDate) -lt 0) { continue }
        if (-not $rankWeight.ContainsKey($row.date)) { $rankWeight[$row.date] = New-Object System.Collections.Generic.List[double] }
        $rankWeight[$row.date].Add([double]$row.weight)
    }
    Write-Host "weight 列を使用: $($rankWeight.Count) 日"
}
$hedgeBeta = @{}
foreach ($sc in $sk.scenarios) {
    if ($null -eq $sc.hedge) { continue }
    $neededCodes = @($neededCodes) + @([string]$sc.hedge.code)
    if ($hedgeBeta.Count -eq 0) { foreach ($row in (Import-Csv (Resolve-ProjectPath $sc.hedge.betaCsv) -Encoding UTF8)) { $hedgeBeta[$row.date] = [double]$row.beta } }
}
$neededCodes = @($neededCodes | Sort-Object -Unique)
Write-Host "loading prices for $($neededCodes.Count) codes that appear in the ranking ..."
$book = Import-PriceBook -Config $config -Codes $neededCodes
$px = $book.px
$calendar = $book.calendar
Write-Host "calendar: $($calendar[0]) - $($calendar[-1]) ($($calendar.Count) days), ranking days: $($ranking.Count)"

function Get-PxIndex {
    param([string]$Code, [string]$Date)
    if (-not $px.ContainsKey($Code)) { return -1 }
    $i = 0
    if ($px[$Code].idx.TryGetValue($Date, [ref]$i)) {
        if ([double]::IsNaN($px[$Code].rawClose[$i])) { return -1 }
        return $i
    }
    return -1
}

# 銘柄ごとに、夜間リターン（前日終値→当日始値）の勝ち数・有効数の累積和（winrate 選別用。Test-WinRateRule.ps1 と同じ計算）
$overnightPrefix = @{}
foreach ($code in $px.Keys) {
    $p = $px[$code]; $n = $p.adjOpen.Count
    $win = [int[]]::new($n + 1); $valid = [int[]]::new($n + 1)
    for ($i = 0; $i -lt $n; $i++) {
        $w = 0; $v = 0
        if ($i -gt 0) {
            $r = $p.adjOpen[$i] / $p.adjClose[$i - 1] - 1
            if (-not [double]::IsNaN($r)) { $v = 1; if ($r -gt 0) { $w = 1 } }
        }
        $win[$i + 1] = $win[$i] + $w; $valid[$i + 1] = $valid[$i] + $v
    }
    $overnightPrefix[$code] = [PSCustomObject]@{ win = $win; valid = $valid }
}

function Select-Codes {
    # シナリオに selection があれば、判断日 $Date までの直近 lookback 日の夜間勝率が高い topK 銘柄に絞る
    # （有効データが lookback の8割未満の銘柄は対象外。同率なら売買代金の順位が上を優先。候補が topK 未満の日は買わない）
    param([string[]]$Codes, [string]$Date, $Selection)
    if ($null -eq $Selection) { return , $Codes }
    $lookback = [int]$Selection.lookback
    $topK = [int]$Selection.topK
    $items = New-Object System.Collections.Generic.List[object]
    for ($k = 0; $k -lt $Codes.Length; $k++) {
        $code = $Codes[$k]
        $i = Get-PxIndex -Code $code -Date $Date
        if ($i -lt 0 -or ($i + 1 - $lookback) -lt 0) { continue }
        $pc = $overnightPrefix[$code]
        $lo = $i + 1 - $lookback
        $hi = $i + 1
        $vCount = $pc.valid[$hi] - $pc.valid[$lo]
        if ($vCount -lt ($lookback * 0.8)) { continue }
        $items.Add([PSCustomObject]@{ code = $code; rank = $k; wr = ($pc.win[$hi] - $pc.win[$lo]) / $vCount })
    }
    if ($items.Count -lt $topK) { return , ([string[]]@()) }
    $picked = @($items | Sort-Object @{ Expression = "wr"; Descending = $true }, @{ Expression = "rank"; Descending = $false } | Select-Object -First $topK | ForEach-Object { $_.code })
    return , ([string[]]$picked)
}

function Get-Allocation {
    # 予算を等分して切り捨て、余りで1株ずつ足す
    #   FillMode "value": 保有額が一番少ない銘柄から足す（既定）
    #   FillMode "rank" : 順位が上の銘柄から順に1株ずつ足す（README 25.）
    #   Weights: 銘柄ごとの予算の比率（合計1に正規化して使う）。省略時は等分
    param([double[]]$LockPrices, [double]$Budget, [string]$FillMode = "value", [double[]]$Weights = $null)
    $n = $LockPrices.Length
    $shares = [int[]]::new($n)
    if ($n -eq 0 -or $Budget -le 0) { return , $shares }
    $w = [double[]]::new($n)
    if ($null -ne $Weights -and $Weights.Length -eq $n) {
        $sum = 0.0; foreach ($x in $Weights) { $sum += $x }
        for ($k = 0; $k -lt $n; $k++) { $w[$k] = $Weights[$k] / $sum }
    } else {
        for ($k = 0; $k -lt $n; $k++) { $w[$k] = 1.0 / $n }
    }
    $remaining = $Budget
    for ($k = 0; $k -lt $n; $k++) {
        $lp = $LockPrices[$k]
        $s = [Math]::Floor($Budget * $w[$k] / $lp)
        $shares[$k] = [int]$s
        $remaining -= $s * $lp
    }
    if ($FillMode -eq "rank") {
        while ($true) {
            $added = $false
            for ($k = 0; $k -lt $n; $k++) {
                if ($LockPrices[$k] -le $remaining) { $shares[$k] = $shares[$k] + 1; $remaining -= $LockPrices[$k]; $added = $true }
            }
            if (-not $added) { break }
        }
    } else {
        while ($true) {
            $best = -1
            $bestVal = [double]::MaxValue
            for ($k = 0; $k -lt $n; $k++) {
                $lp = $LockPrices[$k]
                if ($lp -le $remaining) {
                    $val = $shares[$k] * $lp
                    if ($val -lt $bestVal) { $bestVal = $val; $best = $k }
                }
            }
            if ($best -lt 0) { break }
            $shares[$best] = $shares[$best] + 1
            $remaining -= $LockPrices[$best]
        }
    }
    return , $shares
}

function Test-AtLimit {
    param([double]$Price, [double]$Base)
    $w = Get-LimitWidth -Base $Base
    return ([Math]::Abs($Price - ($Base + $w)) -lt 0.5) -or ([Math]::Abs($Price - ($Base - $w)) -lt 0.5)
}

function Get-PositionsValue {
    # 保有中の銘柄をその日の終値で評価した金額（データが無い日は取得価格で評価）
    param($Positions, [string]$Date)
    $v = 0.0
    foreach ($pos in $Positions) {
        $i = Get-PxIndex -Code $pos.code -Date $Date
        if ($i -ge 0) { $v += $pos.shares * $pos.entryRaw * $px[$pos.code].adjClose[$i] / $pos.entryAdj }
        else { $v += $pos.shares * $pos.entryRaw }
    }
    return $v
}

function Invoke-SKabuSimulation {
    param($Scenario, [double]$CostBps)
    $mode = [string]$Scenario.mode
    if (@("close_to_open_lag1", "open_to_open") -notcontains $mode) { throw "unsupported mode for S株 simulation: $mode" }
    $account = [double]$Scenario.accountCash
    $dailyBudget = [double]$Scenario.dailyBudget
    # compound=true のときは固定の dailyBudget ではなく「注文時点の資産 × budgetRatio」を1日の買付上限にする（利益を再投資）
    $compound = [bool]$Scenario.compound
    $budgetRatio = 1.0
    if ($null -ne $Scenario.budgetRatio) { $budgetRatio = [double]$Scenario.budgetRatio }
    if ($compound) { $dailyBudget = $account * $budgetRatio }
    $cost = $CostBps / 10000.0
    $fillMode = "value"
    if ($Scenario.fillOrder) { $fillMode = [string]$Scenario.fillOrder }

    $cash = $account
    $reserved = 0.0
    $positions = New-Object System.Collections.Generic.List[object]
    $pending = New-Object System.Collections.Generic.List[object]
    $daily = New-Object System.Collections.Generic.List[object]
    $yearly = New-Object System.Collections.Generic.List[object]
    $orders = New-Object System.Collections.Generic.List[object]
    $trades = New-Object System.Collections.Generic.List[object]
    $zeroByCode = @{}

    $counters = [ordered]@{ plannedStocks = 0; boughtStocks = 0; zeroShareStocks = 0; skippedAtLimit = 0; skippedNoData = 0; sellsDelayed = 0; hedgeOrders = 0; hedgeBought = 0; hedgeAmount = 0.0 }
    $yearPnl = 0.0; $yearIdeal = 0.0; $yearInvested = 0.0; $yearTradeDays = 0
    $yearStartEquity = $account
    $curYear = $null
    $peak = $account; $maxDD = 0.0; $maxDDYen = 0.0
    $totalTax = 0.0

    for ($d = 1; $d -lt $calendar.Count; $d++) {
        $date = $calendar[$d]
        $year = $date.Substring(0, 4)
        if ($null -eq $curYear) { $curYear = $year }
        if ($year -ne $curYear) {
            $tax = [Math]::Max(0.0, $yearPnl) * $taxRate
            $cash -= $tax
            $totalTax += $tax
            $yearEndEquity = $daily[$daily.Count - 1].equity - $tax
            if ($yearTradeDays -gt 0 -or $yearPnl -ne 0) { $yearly.Add([PSCustomObject]@{ year = $curYear; trade_days = $yearTradeDays; avg_invested = $(if ($yearTradeDays -gt 0) { $yearInvested / $yearTradeDays } else { 0.0 }); pnl_pretax = $yearPnl; tax = $tax; pnl_aftertax = $yearPnl - $tax; return_on_start_equity = $(if ($yearStartEquity -gt 0) { ($yearPnl - $tax) / $yearStartEquity } else { 0.0 }); ideal_pnl_pretax = $yearIdeal }) }
            $yearStartEquity = $yearEndEquity
            $yearPnl = 0.0; $yearIdeal = 0.0; $yearInvested = 0.0; $yearTradeDays = 0
            $curYear = $year
        }

        $dayPnl = 0.0
        $dayInvested = 0.0

        # 9:00 寄付き: 売り
        $keep = New-Object System.Collections.Generic.List[object]
        foreach ($pos in $positions) {
            if ($pos.sellDay -ne $d) { $keep.Add($pos); continue }
            $i = Get-PxIndex -Code $pos.code -Date $date
            $sold = $false
            if ($i -ge 0) {
                $p = $px[$pos.code]
                $rawOpen = $p.rawOpen[$i]
                $prevIdx = $i - 1
                $stuck = $false
                if ($prevIdx -ge 0 -and $p.rawHigh[$i] -eq $p.rawLow[$i]) { $stuck = Test-AtLimit -Price $rawOpen -Base $p.rawClose[$prevIdx] }
                if (-not $stuck) {
                    $exitAdj = $p.adjOpen[$i]
                    $basis = $pos.shares * $pos.entryRaw
                    $proceeds = $basis * $exitAdj / $pos.entryAdj
                    $fee = $basis * $cost
                    $pnl = $proceeds - $basis - $fee
                    $cash += $proceeds - $fee
                    $yearPnl += $pnl
                    $dayPnl += $pnl
                    $sold = $true
                    $trades.Add([PSCustomObject]@{ code = $pos.code; shares = $pos.shares; buy_date = $pos.buyDate; buy_price = $pos.entryRaw; sell_date = $date; sell_price_raw = $rawOpen; buy_amount = $basis; proceeds_div_adjusted = $proceeds; fee = $fee; pnl = $pnl })
                }
            }
            if (-not $sold) {
                $pos.sellDay = $d + 1
                $counters.sellsDelayed++
                $keep.Add($pos)
            }
        }
        $positions = $keep

        # 9:00 寄付き: open_to_open の買い（前日の引け後に出した注文）
        if ($mode -eq "open_to_open") {
            $stillPending = New-Object System.Collections.Generic.List[object]
            foreach ($o in $pending) {
                if ($o.execDay -ne $d) { $stillPending.Add($o); continue }
                $reserved -= $o.lockAmount
                $i = Get-PxIndex -Code $o.code -Date $date
                if ($i -lt 0) { $counters.skippedNoData++; $o.order.status = "no_data"; continue }
                $p = $px[$o.code]
                $rawOpen = $p.rawOpen[$i]
                if ($p.rawHigh[$i] -eq $p.rawLow[$i] -and (Test-AtLimit -Price $rawOpen -Base $o.base)) { $counters.skippedAtLimit++; $o.order.status = "at_limit"; continue }
                $amount = $o.shares * $rawOpen
                $cash -= $amount
                $dayInvested += $amount
                $counters.boughtStocks++
                $o.order.status = "bought"
                $o.order.fill_price = $rawOpen
                $o.order.amount = $amount
                $positions.Add([PSCustomObject]@{ code = $o.code; shares = $o.shares; entryRaw = $rawOpen; entryAdj = $p.adjOpen[$i]; sellDay = $d + 1; buyDate = $date })
            }
            $pending = $stillPending
        }

        # 10:30〜14:00注文 → 15:30 引け約定: close_to_open_lag1 の買い
        if ($mode -eq "close_to_open_lag1") {
            $t = $calendar[$d - 1]
            if ($ranking.ContainsKey($t)) {
                $codes = Select-Codes -Codes ([string[]]$ranking[$t].ToArray()) -Date $t -Selection $Scenario.selection
                $available = $cash - $reserved
                $budgetCap = $dailyBudget
                if ($compound) { $budgetCap = $budgetRatio * ($cash + (Get-PositionsValue -Positions $positions -Date $t)) }
                $budget = [Math]::Min($budgetCap, $available)
                $bases = [double[]]::new($codes.Length)
                $locks = [double[]]::new($codes.Length)
                for ($k = 0; $k -lt $codes.Length; $k++) {
                    $bi = Get-PxIndex -Code $codes[$k] -Date $t
                    $base = [double]::MaxValue
                    if ($bi -ge 0) { $base = $px[$codes[$k]].rawClose[$bi] }
                    $bases[$k] = $base
                    $locks[$k] = $base + (Get-LimitWidth -Base $base)
                }
                $nStock = $codes.Length
                $stockWeight = 1.0
                $hedgeOk = $true
                if ($null -ne $Scenario.hedge) {
                    # 株の買付額 × β ÷ ETFの倍率 だけインバースETFを買う。予算は株とETFの合計で使う
                    $beta = [double]::NaN
                    if ($hedgeBeta.ContainsKey($t)) { $beta = $hedgeBeta[$t] }
                    $hedgeOk = -not [double]::IsNaN($beta)
                    if ($hedgeOk) { $stockWeight = 1.0 / (1.0 + $beta / [double]$Scenario.hedge.multiple) }
                }
                if (-not $hedgeOk) { $codes = [string[]]@(); $nStock = 0 }
                $wArr = $null
                if ($rankWeight.ContainsKey($t) -and $rankWeight[$t].Count -eq $nStock) { $wArr = [double[]]$rankWeight[$t].ToArray() }
                if ($nStock -gt 0) { $shares = Get-Allocation -LockPrices $locks -Budget ($budget * $stockWeight) -FillMode $fillMode -Weights $wArr }
                else { $shares = [int[]]@() }
                if ($null -ne $Scenario.hedge -and $hedgeOk -and $nStock -gt 0) {
                    $etfCode = [string]$Scenario.hedge.code
                    $ei = Get-PxIndex -Code $etfCode -Date $t
                    $etfBase = [double]::MaxValue
                    if ($ei -ge 0) { $etfBase = $px[$etfCode].rawClose[$ei] }
                    $etfLock = $etfBase + (Get-LimitWidth -Base $etfBase)
                    $stockNotional = 0.0; $stockLock = 0.0
                    for ($k = 0; $k -lt $nStock; $k++) { if ($shares[$k] -gt 0) { $stockNotional += $shares[$k] * $bases[$k]; $stockLock += $shares[$k] * $locks[$k] } }
                    $etfShares = [int][Math]::Floor($stockNotional * $beta / [double]$Scenario.hedge.multiple / $etfBase)
                    $room = [Math]::Floor(($available - $stockLock) / $etfLock)
                    if ($etfShares -gt $room) { $etfShares = [int][Math]::Max(0, $room) }
                    $codes = [string[]](@($codes) + @($etfCode))
                    $bases = [double[]](@($bases) + @($etfBase))
                    $locks = [double[]](@($locks) + @($etfLock))
                    $shares = [int[]](@($shares) + @($etfShares))
                }
                $idealRets = New-Object System.Collections.Generic.List[double]
                $etfIdeal = [double]::NaN
                for ($k = 0; $k -lt $codes.Length; $k++) {
                    $code = $codes[$k]
                    $counters.plannedStocks++
                    $i = Get-PxIndex -Code $code -Date $date
                    if ($i -ge 0 -and ($i + 1) -lt $px[$code].adjOpen.Count) {
                        $idealR = $px[$code].adjOpen[$i + 1] / $px[$code].adjClose[$i] - 1 - $cost
                        if (-not [double]::IsNaN($idealR)) {
                            if ($k -lt $nStock) { $idealRets.Add($idealR) } else { $etfIdeal = $idealR }
                        }
                    }
                    if ($k -ge $nStock) { $counters.hedgeOrders++ }
                    $order = [PSCustomObject]@{ order_date = $date; exec_date = $date; decision_date = $t; code = $code; base_close = $bases[$k]; lock_price = $locks[$k]; budget = $budget; shares = $shares[$k]; lock_amount = $shares[$k] * $locks[$k]; status = ""; fill_price = [double]::NaN; amount = 0.0 }
                    $orders.Add($order)
                    if ($shares[$k] -eq 0) {
                        $order.status = "zero_shares"
                        $counters.zeroShareStocks++
                        if (-not $zeroByCode.ContainsKey($code)) { $zeroByCode[$code] = 0 }
                        $zeroByCode[$code] = $zeroByCode[$code] + 1
                        continue
                    }
                    if ($i -lt 0) { $counters.skippedNoData++; $order.status = "no_data"; continue }
                    $p = $px[$code]
                    $rawClose = $p.rawClose[$i]
                    if (Test-AtLimit -Price $rawClose -Base $bases[$k]) { $counters.skippedAtLimit++; $order.status = "at_limit"; continue }
                    $amount = $shares[$k] * $rawClose
                    $cash -= $amount
                    $dayInvested += $amount
                    $counters.boughtStocks++
                    $order.status = "bought"
                    $order.fill_price = $rawClose
                    $order.amount = $amount
                    if ($k -ge $nStock) { $counters.hedgeBought++; $counters.hedgeAmount += $amount }
                    $positions.Add([PSCustomObject]@{ code = $code; shares = $shares[$k]; entryRaw = $rawClose; entryAdj = $p.adjClose[$i]; sellDay = $d + 1; buyDate = $date })
                }
                if ($idealRets.Count -gt 0) {
                    $sumR = 0.0
                    foreach ($r in $idealRets) { $sumR += $r }
                    $idealBase = $dailyBudget
                    if ($compound) { $idealBase = $budget }
                    $idealDay = $stockWeight * $sumR / $idealRets.Count
                    if ($stockWeight -lt 1.0 -and -not [double]::IsNaN($etfIdeal)) { $idealDay += (1.0 - $stockWeight) * $etfIdeal }
                    $yearIdeal += $idealBase * $idealDay
                }
            }
        }

        # 引け後の注文 → 翌営業日9:00約定: open_to_open の買い
        if ($mode -eq "open_to_open" -and $ranking.ContainsKey($date) -and ($d + 1) -lt $calendar.Count) {
            $codes = Select-Codes -Codes ([string[]]$ranking[$date].ToArray()) -Date $date -Selection $Scenario.selection
            $available = $cash - $reserved
            $budgetCap = $dailyBudget
            if ($compound) { $budgetCap = $budgetRatio * ($cash + (Get-PositionsValue -Positions $positions -Date $date)) }
            $budget = [Math]::Min($budgetCap, $available)
            $bases = [double[]]::new($codes.Length)
            $locks = [double[]]::new($codes.Length)
            for ($k = 0; $k -lt $codes.Length; $k++) {
                $bi = Get-PxIndex -Code $codes[$k] -Date $date
                $base = [double]::MaxValue
                if ($bi -ge 0) { $base = $px[$codes[$k]].rawClose[$bi] }
                $bases[$k] = $base
                $locks[$k] = $base + (Get-LimitWidth -Base $base)
            }
            $shares = Get-Allocation -LockPrices $locks -Budget $budget -FillMode $fillMode
            $nextDate = $calendar[$d + 1]
            $idealRets = New-Object System.Collections.Generic.List[double]
            for ($k = 0; $k -lt $codes.Length; $k++) {
                $code = $codes[$k]
                $counters.plannedStocks++
                $ni = Get-PxIndex -Code $code -Date $nextDate
                if ($ni -ge 0 -and ($ni + 1) -lt $px[$code].adjOpen.Count) {
                    $idealR = $px[$code].adjOpen[$ni + 1] / $px[$code].adjOpen[$ni] - 1 - $cost
                    if (-not [double]::IsNaN($idealR)) { $idealRets.Add($idealR) }
                }
                $order = [PSCustomObject]@{ order_date = $date; exec_date = $nextDate; decision_date = $date; code = $code; base_close = $bases[$k]; lock_price = $locks[$k]; budget = $budget; shares = $shares[$k]; lock_amount = $shares[$k] * $locks[$k]; status = "pending"; fill_price = [double]::NaN; amount = 0.0 }
                $orders.Add($order)
                if ($shares[$k] -eq 0) {
                    $order.status = "zero_shares"
                    $counters.zeroShareStocks++
                    if (-not $zeroByCode.ContainsKey($code)) { $zeroByCode[$code] = 0 }
                    $zeroByCode[$code] = $zeroByCode[$code] + 1
                    continue
                }
                $lockAmount = $shares[$k] * $locks[$k]
                $reserved += $lockAmount
                $pending.Add([PSCustomObject]@{ code = $code; shares = $shares[$k]; lockAmount = $lockAmount; base = $bases[$k]; execDay = $d + 1; order = $order })
            }
            if ($idealRets.Count -gt 0) {
                $sumR = 0.0
                foreach ($r in $idealRets) { $sumR += $r }
                $idealBase = $dailyBudget
                if ($compound) { $idealBase = $budget }
                $yearIdeal += $idealBase * $sumR / $idealRets.Count
            }
        }

        # 引け時点の評価額
        $posValue = 0.0
        foreach ($pos in $positions) {
            $i = Get-PxIndex -Code $pos.code -Date $date
            if ($i -ge 0) { $posValue += $pos.shares * $pos.entryRaw * $px[$pos.code].adjClose[$i] / $pos.entryAdj }
            else { $posValue += $pos.shares * $pos.entryRaw }
        }
        $equity = $cash + $posValue
        if ($equity -gt $peak) { $peak = $equity }
        $dd = ($equity - $peak) / $peak
        if ($dd -lt $maxDD) { $maxDD = $dd; $maxDDYen = $equity - $peak }
        if ($dayInvested -gt 0) { $yearInvested += $dayInvested; $yearTradeDays++ }

        $daily.Add([PSCustomObject]@{ date = $date; cash = $cash; reserved = $reserved; equity = $equity; invested = $dayInvested; positions = $positions.Count; realized_pnl = $dayPnl })
    }

    # 最終年（途中まで）の税金は年末に払うものとして見積もる
    $tax = [Math]::Max(0.0, $yearPnl) * $taxRate
    $totalTax += $tax
    if ($yearTradeDays -gt 0 -or $yearPnl -ne 0) { $yearly.Add([PSCustomObject]@{ year = $curYear; trade_days = $yearTradeDays; avg_invested = $(if ($yearTradeDays -gt 0) { $yearInvested / $yearTradeDays } else { 0.0 }); pnl_pretax = $yearPnl; tax = $tax; pnl_aftertax = $yearPnl - $tax; return_on_start_equity = $(if ($yearStartEquity -gt 0) { ($yearPnl - $tax) / $yearStartEquity } else { 0.0 }); ideal_pnl_pretax = $yearIdeal }) }

    $finalEquity = $daily[$daily.Count - 1].equity - $tax
    $firstTradeDay = $daily | Where-Object { $_.invested -gt 0 } | Select-Object -First 1
    $span = ([datetime]$daily[$daily.Count - 1].date - [datetime]$firstTradeDay.date).TotalDays / 365.25
    $pretaxTotal = 0.0; $idealTotal = 0.0; $investedSum = 0.0; $tradeDays = 0
    foreach ($y in $yearly) { $pretaxTotal += $y.pnl_pretax; $idealTotal += $y.ideal_pnl_pretax; $investedSum += $y.avg_invested * $y.trade_days; $tradeDays += $y.trade_days }

    [PSCustomObject]@{
        scenario        = $Scenario.name
        mode            = $mode
        cost_bps        = $CostBps
        account         = $account
        daily_budget    = $dailyBudget
        compound        = $compound
        budget_ratio    = $budgetRatio
        from            = $firstTradeDay.date
        to              = $daily[$daily.Count - 1].date
        final_equity_aftertax = $finalEquity
        pnl_pretax      = $pretaxTotal
        tax             = $totalTax
        pnl_aftertax    = $pretaxTotal - $totalTax
        cagr_aftertax   = $(if ($span -gt 0 -and $finalEquity -gt 0) { [Math]::Pow($finalEquity / $account, 1 / $span) - 1 } else { 0.0 })
        max_drawdown    = $maxDD
        max_drawdown_yen = $maxDDYen
        trade_days      = $tradeDays
        avg_invested    = $(if ($tradeDays -gt 0) { $investedSum / $tradeDays } else { 0.0 })
        avg_stocks_bought = $(if ($tradeDays -gt 0) { ($counters.boughtStocks - $counters.hedgeBought) / $tradeDays } else { 0.0 })
        hedge_days      = $counters.hedgeBought
        avg_hedge_amount = $(if ($counters.hedgeBought -gt 0) { $counters.hedgeAmount / $counters.hedgeBought } else { 0.0 })
        zero_share_ratio = $(if ($counters.plannedStocks -gt 0) { $counters.zeroShareStocks / $counters.plannedStocks } else { 0.0 })
        skipped_at_limit = $counters.skippedAtLimit
        skipped_no_data = $counters.skippedNoData
        sells_delayed   = $counters.sellsDelayed
        ideal_pnl_pretax = $idealTotal
        yearly          = $yearly
        daily           = $daily
        zero_by_code    = $zeroByCode
        orders          = $orders
        trades          = $trades
    }
}

# --- 実行 ---
$results = New-Object System.Collections.Generic.List[object]
foreach ($sc in $sk.scenarios) {
    foreach ($c in $sk.costBpsList) {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        Write-Host "simulating $($sc.name) cost=${c}bp ..."
        $res = Invoke-SKabuSimulation -Scenario $sc -CostBps ([double]$c)
        Write-Host "  done in $($sw.Elapsed)"
        $results.Add($res)
        $tag = "$($sc.name)_cost$c"
        $res.daily | Export-Csv -Path (Join-Path $outDir "daily_$tag.csv") -NoTypeInformation -Encoding UTF8
        $res.orders | Export-Csv -Path (Join-Path $outDir "orders_$tag.csv") -NoTypeInformation -Encoding UTF8
        $res.trades | Export-Csv -Path (Join-Path $outDir "trades_$tag.csv") -NoTypeInformation -Encoding UTF8
    }
}

$summary = $results | Select-Object scenario, mode, cost_bps, account, daily_budget, compound, budget_ratio, from, to, final_equity_aftertax, pnl_pretax, tax, pnl_aftertax, cagr_aftertax, max_drawdown, max_drawdown_yen, trade_days, avg_invested, avg_stocks_bought, hedge_days, avg_hedge_amount, zero_share_ratio, skipped_at_limit, skipped_no_data, sells_delayed, ideal_pnl_pretax
$summary | Export-Csv -Path (Join-Path $outDir "summary.csv") -NoTypeInformation -Encoding UTF8
$yearRows = foreach ($r in $results) { foreach ($y in $r.yearly) { $y | Select-Object @{N = "scenario"; E = { $r.scenario } }, @{N = "cost_bps"; E = { $r.cost_bps } }, * } }
$yearRows | Export-Csv -Path (Join-Path $outDir "summary_by_year.csv") -NoTypeInformation -Encoding UTF8

$universe = Import-Csv -Path (Resolve-ProjectPath $config.universe.csv) -Encoding UTF8
$nameByCode = @{}
foreach ($u in $universe) { $nameByCode[$u.code] = $u.name }

Write-Host ""
Write-Host "=== S株シミュレーション ($($results[0].from) - $($results[0].to)) ==="
foreach ($r in $results) {
    Write-Host ""
    if ($r.compound) {
        Write-Host ("[{0}] コスト{1}bp / 元手{2:N0}円 / 1日の買付上限は資産の{3:P0}（利益を再投資）" -f $r.scenario, $r.cost_bps, $r.account, $r.budget_ratio)
    } else {
        Write-Host ("[{0}] コスト{1}bp / 口座{2:N0}円 / 1日の買付上限{3:N0}円" -f $r.scenario, $r.cost_bps, $r.account, $r.daily_budget)
    }
    Write-Host ("  最終資産(税引後) {0,12:N0}円  損益 税前 {1,10:N0}円 / 税 {2,9:N0}円 / 税後 {3,10:N0}円  年率(税後) {4,6:N1}%" -f $r.final_equity_aftertax, $r.pnl_pretax, $r.tax, $r.pnl_aftertax, ($r.cagr_aftertax * 100))
    Write-Host ("  最大ドローダウン {0,6:N1}% ({1:N0}円)   理想(予算を上位{2}に端数なしで等分)の税前損益 {3,10:N0}円 → 実現率 {4:P0}" -f ($r.max_drawdown * 100), $r.max_drawdown_yen, $topN, $r.ideal_pnl_pretax, $(if ($r.ideal_pnl_pretax -ne 0) { $r.pnl_pretax / $r.ideal_pnl_pretax } else { 0 }))
    Write-Host ("  取引日 {0}日  平均買付額 {1:N0}円  平均買付銘柄数 {2:N1}  0株になった枠 {3:P1}  値幅張り付きで不約定 {4}件  データ無し {5}件  売り持ち越し {6}件" -f $r.trade_days, $r.avg_invested, $r.avg_stocks_bought, $r.zero_share_ratio, $r.skipped_at_limit, $r.skipped_no_data, $r.sells_delayed)
    if ($r.hedge_days -gt 0) { Write-Host ("  インバースETFを買った日 {0}日  平均買付額 {1:N0}円（平均買付額・税・損益はETFを含む）" -f $r.hedge_days, $r.avg_hedge_amount) }
    $z = $r.zero_by_code.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 5 | ForEach-Object { "{0}({1}) {2}回" -f $nameByCode[$_.Key], $_.Key, $_.Value }
    if ($z) { Write-Host "  0株になった回数が多い銘柄: $($z -join ' / ')" }
    Write-Host ("  {0,-6} {1,6} {2,12} {3,12} {4,10} {5,12} {6,10} {7,12}" -f "年", "取引日", "平均買付額", "税前損益", "税", "税後損益", "期初比", "理想税前")
    foreach ($y in $r.yearly) {
        Write-Host ("  {0,-6} {1,6} {2,12:N0} {3,12:N0} {4,10:N0} {5,12:N0} {6,9:N1}% {7,12:N0}" -f $y.year, $y.trade_days, $y.avg_invested, $y.pnl_pretax, $y.tax, $y.pnl_aftertax, ($y.return_on_start_equity * 100), $y.ideal_pnl_pretax)
    }
}

# 資産推移のチャート
Add-Type -AssemblyName System.Windows.Forms.DataVisualization
$chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
$chart.Width = 1100
$chart.Height = 550
$chart.ChartAreas.Add((New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea))
$chart.Legends.Add((New-Object System.Windows.Forms.DataVisualization.Charting.Legend)) | Out-Null
foreach ($r in $results) {
    $s = New-Object System.Windows.Forms.DataVisualization.Charting.Series
    $s.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
    $s.Name = "$($r.scenario) cost$($r.cost_bps)bp (pre-tax equity)"
    $s.BorderWidth = 2
    foreach ($row in $r.daily) { $s.Points.AddXY($row.date, $row.equity) | Out-Null }
    $chart.Series.Add($s)
}
$title = New-Object System.Windows.Forms.DataVisualization.Charting.Title
$title.Text = "S-kabu simulation: account equity (yen, before tax)"
$chart.Titles.Add($title)
$chartPath = Join-Path $outDir "equity.png"
$chart.SaveImage($chartPath, [System.Windows.Forms.DataVisualization.Charting.ChartImageFormat]::Png)

Write-Host ""
Write-Host "saved $(Join-Path $outDir 'summary.csv'), summary_by_year.csv, daily_*.csv, equity.png"
