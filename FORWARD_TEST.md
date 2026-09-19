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

## 2-2. 自動で回す（タスクスケジューラ）

上の「毎営業日」を手で叩く代わりに、Windows のタスクスケジューラに登録できる。管理者権限は不要。

```powershell
powershell -File ps\Register-ScheduledTasks.ps1            # 登録（17時=更新、22時=確認）
powershell -File ps\Register-ScheduledTasks.ps1 -Show      # 状態と次回実行を見る
powershell -File ps\Register-ScheduledTasks.ps1 -Unregister # 取り消す
```

| 時刻 | 中身 |
|---|---|
| 17:00 | `Invoke-DailyUpdate.ps1` — 株価と指数を取り直し、ランキング・B2の選定・6通りのシミュレーション・カレンダー・翌営業日の銘柄までを一度に更新する（40〜50分） |
| 22:00 | `Test-DailyUpdate.ps1` — 当日のデータがそろっているか確かめ、駄目なら取り直す |

- **休業日は両方とも何もしない。** 土日と祝日は `ps\Common.ps1` の `$script:JpxHolidays` で判定する。年をまたぐ前にこのリストを足すこと（範囲外の日付を渡すと警告が出る）。
- 17時の更新は `Fetch-UniversePrices.ps1 -Refresh` を使う。1銘柄ずつ `.tmp` に書いてから置き換えるので、途中で失敗しても失敗した銘柄は前のファイルが残る（CSVをまとめて消す必要はない）。
- 16時より前に `Invoke-DailyUpdate.ps1` を実行すると、取引時間中の値をつかまないように止まる。手で動かすときは `-Force`。
- ログは `logs\daily_YYYYMMDD.log` と `logs\check_YYYYMMDD.log`。結果は `logs\last_run.json` / `logs\last_check.json` に残る（`logs\` は追跡しない）。
- 22時の確認は、N225の最終日・銘柄CSVの更新時刻・抜き取りした銘柄の最終行・`today_picks.json` の `asOf` を見る。過去の日を調べ直すときは `-AsOf 2026-09-18`、取り直しをさせたくないときは `-NoRepair`。
- **PCがスリープ・電源断のときは動かない。** `StartWhenAvailable` を入れてあるので、起動後に取りこぼした回をできるだけ早く実行する。ログオンしている間だけ動く設定（パスワードを預けずに済ませるため）。


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
- `DATA` は `reports\b3l_2026\calendar_data.json` の中身で置き換える（シミュレーション結果から `ps\Export-CalendarData.ps1` が作る）

この置き換えは `powershell -File ps\Update-Calendar.ps1` が行う（置き換え前の版は `calendar.html.bak` に残る）。
株価の更新からカレンダーまで一度にやるなら `powershell -File ps\Update-B3L2026.ps1`。

## 注意

- 株価データは**2022年1月以降**。B2は250営業日の助走が必要なので、このデータで計算できるのは2023年以降。
- 現在プライムに上場している銘柄だけのデータなので、生存者バイアスがある。
- 表示・計算はすべて過去データのシミュレーション。投資判断の助言ではない。
