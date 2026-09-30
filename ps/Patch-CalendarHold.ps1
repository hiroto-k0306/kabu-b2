# web/calendar.html を「放置」パターン（calendar_data.json の variants に kind = "hold" があるもの）に対応させる。
# 1回だけ実行すればよい（もう直してあれば何もしない）。const DATA / const TODAY の行には触れない。
#   - 日ごとの表示の「買」「売」を、放置のときは「前」「後」（前の営業日・この日の引けの評価額）にする
#   - 日付を選んだときの見出し・説明・買付額/売却額の名前を、放置のときは評価額の言い方にする
#   - 画面上部の条件の説明と、下の注意書きに、放置パターンの説明を足す
# 使い方: powershell -File ps\Patch-CalendarHold.ps1
param([string]$Html = "web/calendar.html")

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"
Set-Location (Get-ProjectRoot)

$path = Resolve-ProjectPath $Html
if (-not (Test-Path $path)) { throw "$Html がない" }
$text = [IO.File]::ReadAllText($path)
if ($text.Contains("const isHold")) { Write-Host "already patched: $Html"; exit 0 }
$nl = if ($text.Contains("`r`n")) { "`r`n" } else { "`n" }

function Set-Once {
    param([string]$Old, [string]$New)
    $n = [regex]::Matches($script:text, [regex]::Escape($Old)).Count
    if ($n -ne 1) { throw "置き換え元が $n 個ある（1個のはず）: $Old" }
    $script:text = $script:text.Replace($Old, $New)
}

# 1. 詳細欄の見出しに id を付ける（放置のとき言い方を変えるため）
Set-Once '<dt>買付額（引け）</dt>' '<dt id="dBuyLabel">買付額（引け）</dt>'
Set-Once '<dt>売却額（翌寄り）</dt>' '<dt id="dSellLabel">売却額（翌寄り）</dt>'

# 2. 放置パターンかどうか
Set-Once 'const labelOf = k => DATA.variants[k].label || k;' ('const labelOf = k => DATA.variants[k].label || k;' + $nl + '  const isHold = k => !!(DATA.variants[k] && DATA.variants[k].kind === "hold");')

# 3. カレンダーの日ごとの「買」「売」
Set-Once '<em>買</em> '' + yen(rec.buyAmount)' '<em>'' + (isHold(curVariant) ? "前" : "買") + ''</em> '' + yen(rec.buyAmount)'
Set-Once '<em>売</em> '' + yen(rec.sellAmount)' '<em>'' + (isHold(curVariant) ? "後" : "売") + ''</em> '' + yen(rec.sellAmount)'

# 4. 日付を選んだときの見出し・説明
$ind = "    "
Set-Once 'document.getElementById("detailTitle").textContent = fmt(iso) + " に買った" + rec.trades.length + "銘柄";' (
    'const hold = isHold(curVariant);' + $nl + $ind +
    'document.getElementById("detailTitle").textContent = hold ? fmt(iso) + " の保有（" + rec.trades.length + "銘柄）" : fmt(iso) + " に買った" + rec.trades.length + "銘柄";')
Set-Once 'document.getElementById("detailSub").textContent = fmt(iso) + "の引けで買い → " + fmt(rec.sellDate) + "の寄りで売却";' (
    'document.getElementById("detailSub").textContent = hold ? fmt(iso) + "の引けの評価額 → " + fmt(rec.sellDate) + "の引けの評価額（放置・売買なし）" : fmt(iso) + "の引けで買い → " + fmt(rec.sellDate) + "の寄りで売却";' + $nl + $ind +
    'document.getElementById("dBuyLabel").textContent = hold ? "評価額（この日の引け）" : "買付額（引け）";' + $nl + $ind +
    'document.getElementById("dSellLabel").textContent = hold ? "評価額（次の営業日の引け）" : "売却額（翌寄り）";')

# 5. 画面上部の条件の説明
Set-Once '円・利益を再投資 / S株・引け買い → 翌寄り売り";' '円・利益を再投資 / " + (isHold(curVariant) ? "放置（最初の買い日の引けに全額投入して売買しない。東証ETFで代用・分配金込み）" : "S株・引け買い → 翌寄り売り");'

# 6. 注意書き
Set-Once '<span id="taxNote"></span>投資判断の助言ではありません。' '<span id="taxNote"></span>「放置」のS&P500・オルカンは、最初の買い日の引けに元手を全額投入（金額指定・端数なし）して、その後は売買しない場合です。投資信託の基準価額は取れないので、連動する東証ETF（2558・2559）の分配金込みの価格で代用しており、信託報酬・為替・ETFの価格の乖離は投資信託と少し違います。取得元の異常値の日は除いています。投資判断の助言ではありません。'

# 7. カレンダー見出しの説明（買・売 → 前・後）と、「次の営業日に買う銘柄」欄（放置は売買しない）
Set-Once '<span class="sub">買＝引けで買った金額、売＝翌営業日の寄りで売った金額（円）</span>' '<span class="sub" id="calLegend">買＝引けで買った金額、売＝翌営業日の寄りで売った金額（円）</span>'
Set-Once 'document.getElementById("calTitle").textContent = Y + "年" + M + "月（" + labelOf(curVariant) + "）";' (
    'document.getElementById("calTitle").textContent = Y + "年" + M + "月（" + labelOf(curVariant) + "）";' + $nl + $ind +
    'document.getElementById("calLegend").textContent = isHold(curVariant) ? "前＝この日の引けの評価額、後＝次の営業日の引けの評価額（円・売買なし）" : "買＝引けで買った金額、売＝翌営業日の寄りで売った金額（円）";')
Set-Once 'const rowsEl = document.getElementById("todayRows"), footEl = document.getElementById("todayFoot");' (
    'const rowsEl = document.getElementById("todayRows"), footEl = document.getElementById("todayFoot");' + $nl + $ind +
    'if (isHold(curVariant)) {' + $nl + $ind +
    '  document.getElementById("todayTitle").textContent = "次の営業日の売買（" + labelOf(curVariant) + "）";' + $nl + $ind +
    '  document.getElementById("todaySub").textContent = "放置のため売買はありません（最初の買い日に全額投入したまま保有）";' + $nl + $ind +
    '  document.getElementById("sched").innerHTML = ""; rowsEl.innerHTML = ""; footEl.innerHTML = "";' + $nl + $ind +
    '  return;' + $nl + $ind +
    '}')
[IO.File]::WriteAllText($path, $text, (New-Object Text.UTF8Encoding $false))
Write-Host "patched $Html"
