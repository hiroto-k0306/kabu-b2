# kabu-b2 — B2モデルの実行に必要な資産

B2モデル（夜間の勢い×値動きの小ささ）を動かすために必要なものだけを集めたフォルダ。
検討の経緯は [MODEL_HISTORY.md](MODEL_HISTORY.md)、研究に使った他のスクリプト・生データ一式は元プロジェクト `C:\Users\user\Desktop\kabu` にある。

2026-11-04に、将来の約定可否を銘柄選択から分離し、売却日に実現済みの目的変数だけをボラティリティモデルのラグ・学習・再補正へ使う修正を統合した。未約定の結果は欠損として保持する。データ欠損日の保有評価は当日までの最終価格を使い、フォワード成績のpnl表示は「手数料控除後・税引前」とする。既存の成績ファイルは再計算していないため、この修正後の成績として扱わない。

`pwsh -NoProfile -File ps/Test-TemporalBoundaries.ps1` は合成データだけで時点境界を検証する。外部取得・実取引は行わない。[統合記録](reports/temporal-boundary-integration-20261104.md)に範囲と残件を記載した。

## 中身

```
ps/
  Common.ps1                    共通処理（設定の読み込み、株価の読み込み・検証、Yahooからの取得）
  CrossSectionStudy.cs          全銘柄×長期の集計エンジン（Add-TypeでC#としてコンパイル）
  Test-CrossSectionSignals.ps1  信号の検証と、日ごとの上位銘柄の出力
  Simulate-SKabu.ps1            S株での株数・余力・値幅制限・税金のシミュレーション
  Fetch-UniversePrices.ps1      銘柄ごとの日足の取得（続きから再開できる）
  Fetch-MarketData.ps1          日経平均・ETF・S&P500・VIX・ドル円・10年国債利回りの取得
  Import-JpxListing.ps1         JPXの上場銘柄一覧(xlsx)から銘柄リストを作る
  config.prime_long.json        株価取得の設定（2000年以降。2026-09-24時点の全件は kabuData リポジトリに保存）
  config.prime_daily.json       日次更新の株価取得の設定（2020年以降）
  config.b3l.json               S株シミュレーション（10銘柄・2022年開始）
  config.b3l_top5.json          同（5銘柄）
  config.b3l_2026_top3/5/10.json 同（2026年開始・再投資あり）
data/
  raw/universe/prime.csv        東証プライム1,556銘柄と33業種（2026年8月のJPX一覧より）
  raw/market/daily_since2000/N225.csv  日経平均の日足（営業日カレンダーとして使う）
  processed/b3l/*.csv           B2が選んだ日ごとの銘柄（10/5/3/1銘柄、2022-01〜2026-09）
  prices_zip/prime_since2022_part1〜4.zip  プライム1,556銘柄の日足（2022-01以降）。展開して使う
reports/b3l_2026/calendar_data.json  カレンダー用のデータ（3・5・10銘柄、2026年）
web/calendar.html               日別収支のカレンダー（ブラウザで開く。データは埋め込み済み）
MODEL_HISTORY.md                v1から現在までの検討の経緯
```

## 使い方

### 1. 株価データを展開する

```powershell
cd C:\Users\user\Desktop\kabu-b2
New-Item -ItemType Directory -Force data\raw\stocks\prime_since2000 | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem
Get-ChildItem data\prices_zip\*.zip | ForEach-Object {
  [IO.Compression.ZipFile]::ExtractToDirectory($_.FullName, (Resolve-Path data\raw\stocks\prime_since2000).Path)
}
```

展開すると1,556ファイル・約220MBになる。ファイル名・列はスクリプトが期待する形（`data/raw/stocks/prime_since2000/<code>.csv`）のまま。

### 2. B2の上位銘柄を出す

```powershell
powershell -File ps\Test-CrossSectionSignals.ps1 -Phase Reference -Signal B3L            # 10銘柄
powershell -File ps\Test-CrossSectionSignals.ps1 -Phase Reference -Signal B3L -TopK 5    # 5銘柄
```

`reports/cross_section/daily_Reference_minprice0.csv` に日ごとの上位銘柄（`picks_B3L`）とリターンが出る。
S株シミュレーションに渡すには、この列を `date,rank,code` の形に変換して `data/processed/b3l/` に置く（すでに変換済みのファイルが入っている）。

### 3. S株シミュレーション

```powershell
$env:KABU_CONFIG = "ps/config.b3l.json"; powershell -File ps\Simulate-SKabu.ps1
```

### 4. データを最新にする

```powershell
$env:KABU_CONFIG = "ps/config.prime_long.json"; powershell -File ps\Fetch-UniversePrices.ps1  # 既にあるファイルは飛ばす
powershell -File ps\Fetch-MarketData.ps1
```

最新にする場合は、更新したい銘柄のCSVを削除してから実行する（あるファイルは取得しない作りのため）。

## 前提と注意

- Windows標準のPowerShell 5.1と.NETだけで動く（追加のインストールは不要）。`.ps1`はUTF-8（BOM付き）で保存する
- 株価は `open/high/low/close` が分割・配当調整済み、`raw_*` が当時の実際の価格。株数や金額は `raw_*`、リターンは調整済みを使う
- 価格データは**2022-01以降に絞ってある**。B2は250日の助走が必要なため、このデータで計算できるのは概ね2023年以降。それより前を扱う場合は元プロジェクトの `data/raw/stocks/prime_since2000`（2000年以降、約1.1GB）を使う
- 現在プライムに上場している銘柄だけのデータなので、生存者バイアスがある（上場廃止銘柄は含まれない）
- 実際の売買の記録ではなく、過去データのシミュレーション。投資判断の助言ではない


## 日次更新のデータ欠損・朝の補完

- `Invoke-DailyUpdate.ps1` は開始時に日本時間で確定済みの対象営業日を決め、取得・検査の全段階に `-AsOf yyyy-MM-dd` を渡す。17時前は前営業日、17時以降は当営業日が対象。`-Force` でも未確定日は指定できない。
- Yahoo がHTTP成功を返しても、対象日の有効な日足がなければ待機して再取得する。既定は最大3回、再試行前に60秒・120秒待つ。個別株は欠損した銘柄だけ再試行し、成功済み銘柄は再取得しない。
- 対象日が欠ける、空、無効なOHLCの場合は既存CSVを保持する。保存は一時ファイルから置き換える。取得の対象日・最新日・試行回数・失敗理由は `logs/stock_fetch.json` と `logs/market_fetch.json` に残る。
- 朝の起動が寄り付き後に遅れても、前営業日までのデータだけを保存して補完できる。取得要求の終了日時に加え、戻った日足も対象日以下に制限する。米国指数・為替は日本の対象日より前の暦日までに制限し、7暦日超古い系列は失敗扱いにする。これは米国・為替の厳密な休日別の鮮度保証ではない。
- 集計前に `Test-DailyUpdate.ps1 -NoRepair -PricesOnly -AsOf ...` を実行する。元の更新時刻・抜き取り検査を残し、期待する銘柄一覧を分母にした有効な対象日データ98%以上、日本の指数・ETFの対象日一致、対象日より新しいCSVがないことも要求する。不合格なら集計・公開用commitに進まない。
- `pwsh -NoProfile -File ps/Test-DataFreshness.ps1` はHTTP・待機をモックした合成テスト。価格取得、本番データ変更、新規バックテスト、実取引は行わない。Windows PowerShell 5.1でも実行可能。既存の時点境界検査は `pwsh -NoProfile -File ps/Test-TemporalBoundaries.ps1 -SkipModel` でモデル学習なしに実行できる。
