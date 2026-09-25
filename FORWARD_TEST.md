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
powershell -File ps\Register-ScheduledTasks.ps1            # 登録（17時=更新、翌朝8時=確認）
powershell -File ps\Register-ScheduledTasks.ps1 -Show      # 状態と次回実行を見る
powershell -File ps\Register-ScheduledTasks.ps1 -Unregister # 取り消す
```

| 時刻 | 中身 |
|---|---|
| 17:00 | `Invoke-DailyUpdate.ps1` — 株価と指数を取り直し、ランキング・B2の選定・6通りのシミュレーション・カレンダー・翌営業日の銘柄までを一度に更新する（40〜50分） |
| 翌朝 08:00 | `Test-DailyUpdate.ps1` — **直近の営業日**のデータがそろっているか確かめ、駄目なら取り直す |

- **17時の更新は休業日には何もしない。** 朝の確認は休業日でも動くが、見るのは直近の営業日なので、
  土日のうちは同じ日（金曜の分）を繰り返し確かめることになる。
- 土日と祝日は `ps\Common.ps1` の `$script:JpxHolidays` で判定する。年をまたぐ前にこのリストを足すこと（範囲外の日付を渡すと警告が出る）。
- 17時の更新は `Fetch-UniversePrices.ps1 -Refresh`（`config.prime_daily.json`、2020年以降）を使う。1銘柄ずつ `.tmp` に書いてから置き換えるので、途中で失敗しても失敗した銘柄は前のファイルが残る（CSVをまとめて消す必要はない）。
- 16時より前に `Invoke-DailyUpdate.ps1` を実行すると、取引時間中の値をつかまないように止まる。手で動かすときは `-Force`。
- ログは `logs\daily_YYYYMMDD.log` と `logs\check_YYYYMMDD.log`。結果は `logs\last_run.json` / `logs\last_check.json` に残る（`logs\` は追跡しない）。
- 朝の確認が見るのは**直近の営業日**（金曜の分は土曜の朝、木曜の分が金曜が祝日なら金曜の朝）。17時以降に手で動かしたときだけ当日を見る。過去の日を調べ直すときは `-AsOf 2026-09-18`、取り直しをさせたくないときは `-NoRepair`。
- 確認する中身は、N225の最終日・銘柄CSVの更新時刻・抜き取りした銘柄の最終行・`today_picks.json` の `asOf`。
- **取引時間中（9:00〜16:00）は取り直さない。** 途中の値をつかむため、17時の更新に任せる（急ぐときだけ `-Force`）。朝8時は寄り付き前なので問題ない。
- 休業日の17時は何もせず、`last_run.json` も書き換えない（直近に実際に動いた回が分かるようにするため）。
- **PCがスリープ・電源断のときは動かない。** `StartWhenAvailable` を入れてあるので、起動後に取りこぼした回をできるだけ早く実行する。ログオンしている間だけ動く設定（パスワードを預けずに済ませるため）。

## 2-3. 自動で回す（GitHub Actions）

PCを使わず、GitHub のサーバーで同じ処理を回す。`.github/workflows/daily-update.yml` が `main` にあれば動く。
**タスクスケジューラと二重に動かさないこと**（`powershell -File ps\Register-ScheduledTasks.ps1 -Unregister`）。

| 時刻（日本時間） | 中身 |
|---|---|
| 平日 17:17 | `Test-NeedsUpdate.ps1` で当日の分が git に無いと分かったら、zip を展開して `Invoke-DailyUpdate.ps1 -Force` → `Test-DailyUpdate.ps1 -NoRepair` → `main` に commit |
| 毎日 7:43 | 取りこぼしの補完。直近の営業日の分が無いときだけ上と同じことをする |

- 手元では `git pull` するだけで `web\calendar.html` と `reports\today_picks.json` が最新になる。
- 毎回まっさらな環境なので、zip（2022年以降）を展開したうえで全銘柄を2020年から取り直す（1銘柄ずつ）。全体で30分前後。
  2022年からのB2には250営業日の助走が要るので2020年からにしてある。2000年から取った場合と2022年以降の結果が同じことは確認済み。
  売買代金ランキングと比較用ユニバース（`data/processed/`）は2020年4月からになる（2000〜2019年分は git の履歴と kabuData に残る）。
- 休業日かどうかの判定と取引時間中に動かない決まりは、タスクスケジューラ版と同じ（`Common.ps1` の祝日リストを使う）。
- 失敗すると GitHub から失敗通知のメールが届く。ログは実行結果の Artifacts（`daily-logs`、30日保存）。
- 手で動かすときは GitHub の Actions → daily-update → Run workflow。最新でも取り直すなら `force` にチェック。
- GitHub の都合で定時から数十分遅れることがある。取引時間中（9:00〜16:00）にずれ込んだ回は何もしない。
- Linux の PowerShell 7 で動くので、`equity.png`（Windows のみのチャート）は作らない。また小数の桁数が
  Windows（15桁）と違う（最大17桁）ため、初回は `data/processed` などのCSVが全行書き換わる。値は同じ。


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
| `config.prime_long.json` | 株価取得用（2000年以降。新しいモデルを作るとき用） |
| `config.prime_daily.json` | 日次更新の株価取得用（2020年以降。`Invoke-DailyUpdate.ps1` が使う） |
| `config.prime.json` / `config.nikkei_pit.json` | 売買代金ランキング作成用 |

## 4. カレンダーを更新したいとき

`web\calendar.html` の `const DATA = {...}` と `const TODAY = {...}` が埋め込みデータ。

- `TODAY` は `reports\today_picks.json` の中身で置き換える
- `DATA` は `reports\b3l_2026\calendar_data.json` の中身で置き換える（シミュレーション結果から `ps\Export-CalendarData.ps1` が作る）

この置き換えは `powershell -File ps\Update-Calendar.ps1` が行う（置き換え前の版は `calendar.html.bak` に残る）。
株価の更新からカレンダーまで一度にやるなら `powershell -File ps\Update-B3L2026.ps1`。

## 5. 未知データでの成績を残す

モデルを固めた後に出てきたデータでの成績は `FORWARD_TEST_RESULTS.md` に積み上がる。
元になる台帳は `reports\forward_test\daily.csv`（1行=1日1系統）と `trades.csv`（1行=1建玉）。

```powershell
powershell -File ps\Update-ForwardTest.ps1              # 台帳に足して読み物を作り直す
powershell -File ps\Update-ForwardTest.ps1 -Rebuild     # 台帳ごと作り直す（普段は使わない）
```

- `Update-B3L2026.ps1` の最後で呼ばれるので、17時の自動更新に含まれる。
- **一度書いた行は書き換えない。** 毎日シミュレーションをやり直すと、分割・配当で過去の
  調整済み価格が変わったときに昔の数字まで動いてしまい、「その時どうだったか」の記録に
  ならないため。まだ無い日付だけを足す。何度実行しても増えない。
- 区切りは **2026-09-17に買った分から**（`-StartDate` で変えられる）。配布時点のカレンダーが
  持っていた最後の建玉が9/16買いなので、その次の営業日から先が未知データにあたる。


## 注意

- 株価データは**2022年1月以降**。B2は250営業日の助走が必要なので、このデータで計算できるのは2023年以降。
- 現在プライムに上場している銘柄だけのデータなので、生存者バイアスがある。
- 表示・計算はすべて過去データのシミュレーション。投資判断の助言ではない。
