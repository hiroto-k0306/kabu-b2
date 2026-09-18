# 別端末での運用手順（未知のデータでの検証用）

このフォルダ一式をダウンロードすれば、別のWindows端末でそのまま使える。追加インストールは不要（Windows標準のPowerShell 5.1 と .NET だけ）。

## 1. 初回だけ

```powershell
cd <このフォルダ>
# 株価データ（2022年以降、1,556銘柄）を展開する
New-Item -ItemType Directory -Force data\raw\stocks\prime_since2000 | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem
Get-ChildItem data\prices_zip\*.zip | ForEach-Object {
  [IO.Compression.ZipFile]::ExtractToDirectory($_.FullName, (Resolve-Path data\raw\stocks\prime_since2000).Path)
}
```

展開後は約220MB・1,556ファイル。`web\calendar.html` はブラウザで開けばそのまま見られる（データは埋め込み済み）。

## 2. 毎営業日（大引け後、16時以降）

```powershell
# 1) 株価を最新にする（既にあるファイルは取得しないので、更新したい銘柄は削除してから実行）
$env:KABU_CONFIG = "ps/config.prime_long.json"; powershell -File ps\Fetch-UniversePrices.ps1
powershell -File ps\Fetch-MarketData.ps1

# 2) 売買代金ランキングを作り直す（日経225の各時点構成 と 東証プライム）
$env:KABU_CONFIG = "ps/config.nikkei_pit.json"; powershell -File ps\Build-TurnoverRanking.ps1
$env:KABU_CONFIG = "ps/config.prime.json";      powershell -File ps\Build-TurnoverRanking.ps1

# 3) 次の営業日に買う銘柄を出す（画面表示＋ reports\today_picks.json）
powershell -File ps\Get-TodayPicks.ps1 -Budget 500000
```

- **取得は大引け後に。** 取引時間中（9:00〜15:00）に取得すると最新行が途中の値になり、判断が狂う。
- `Get-TodayPicks.ps1` は買い日・売り日（休業日を考慮）も表示する。祝日は内蔵リストなので取引所のカレンダーで確認すること。
- 注文は **買い: 次の営業日の10:30〜14:00（当日15:30の終値で約定）→ 売り: その日の14:00〜翌営業日7:00（翌営業日9:00の始値で約定）**。

### 株価を毎日全部取り直すと時間がかかる場合

`Fetch-UniversePrices.ps1` は「ファイルがあれば取得しない」作りなので、全銘柄を更新するには既存CSVを消す必要があり、約25分かかる。検証だけが目的なら、毎日取得せず数日〜1週間おきにまとめて取得しても、各判断日の結果は同じように再現できる（ただし分割・配当で過去の調整済み価格が変わると、当時の判断と少しずれる可能性がある）。

## 3. 結果を突き合わせる

```powershell
# 指定した元手・開始日でS株シミュレーション（正確な株数・余力・税金つき）
$env:KABU_CONFIG = "ps/config.b3l.json"; powershell -File ps\Simulate-SKabu.ps1
```

`ps\config.*.json` の `accountCash` / `dailyBudget` / `startDate` / `costBpsList` を編集して使う。

| 設定ファイル | 内容 |
|---|---|
| `config.b3l.json` | B2 10銘柄（2022年開始） |
| `config.b3l_top5.json` / `config.b3l_top3.json` | B2 5銘柄 / 3銘柄 |
| `config.b3l_2026_top3/5/10.json` | B2 3/5/10銘柄（2026年開始・利益を再投資） |
| `config.prime_long.json` | 株価取得用（2000年以降） |
| `config.prime.json` / `config.nikkei_pit.json` | 売買代金ランキング作成用 |

## 4. カレンダーを更新したいとき

`web\calendar.html` の `const DATA = {...}` と `const TODAY = {...}` が埋め込みデータ。

- `TODAY` は `reports\today_picks.json` の中身で置き換える
- `DATA` は `reports\b3l_2026\calendar_data.json` の中身で置き換える（シミュレーション結果から作るファイル）

## 注意

- 株価データは**2022年1月以降**。B2は250営業日の助走が必要なので、このデータで計算できるのは2023年以降。
- 現在プライムに上場している銘柄だけのデータなので、生存者バイアスがある。
- 表示・計算はすべて過去データのシミュレーション。投資判断の助言ではない。
