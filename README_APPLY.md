# kabu-b2 への反映パッチ（B3・M0・T2 の前向き検証 と カレンダーの「放置」パターン）

作成日 2026-09-30。このフォルダのファイルを、kabu-b2 のクローンの同じ場所に上書きコピーして、commit・push する。
私（Claude）の環境には git も GitHub の認証も無いので、リポジトリへの反映（commit・push）はご自身で行う。

## 何を足したか

### A. 未知データの台帳（`reports/forward_test/daily.csv`）に3系統を足す（カレンダーには出さない）

| 系統 | 中身 | 開始（買い日） |
|---|---|---|
| `B3 10銘柄` | 低ボラの絞り込みをしない B3 の上位10銘柄 | 2026-09-24 |
| `M0 (B3とB2の混合)` | B3 の上位10 と B2 の上位10の50:50混合（重複する銘柄は重み2、20銘柄前後） | 2026-09-24 |
| `T2 (二段階)` | 上の約20銘柄を「直近20営業日の夜間リターン平均」で並べ替えた上位10銘柄 | 2026-09-24 |

- 開始日は、定義を 2026-09-17 までのデータで固めたため、9/18 判断（買い日は休業日をはさんで 9/24）以降。窓調整の系統（9/28開始）と同じ考え方。
- 選定は判断日の引けまでのデータだけを使い、買い日の引けで買って翌営業日の寄りで売る（B2 と同じ）。S株シミュレーション（1株単位、値幅制限、コスト3bp、再投資）も B2 と同じ `Simulate-SKabu.ps1` で行う。
- 定義の出どころ: 私の側の検証（kabu2 の README「変更15〜25」）。T2 の第2段は、事前に登録した「直近20営業日の夜間リターン平均」。

### B. カレンダーに「放置」パターン（S&P500・オルカン）を2つ足す

- カレンダーの最初の買い日（2026-01-06）の引けに元手50万円を全額投入し、以後は売買しない場合。
- 投資信託の基準価額は日足の取得先に無いので、連動する東証ETF（2558 MAXIS 米国株式(S&P500)上場投信、2559 MAXIS 全世界株式(オール・カントリー)上場投信）の**分配金込みの価格**で代用する。金額指定で全額投入（端数なし）。
- 日ごとの損益は「その日の引けの評価額 → 次の営業日の引けの評価額」。カレンダー右上の元手・開始日の入力も、他の戦略と同じ仕組みで効く（開始日を変えると、その日の引けに投入した場合になる）。
- **取得元（Yahoo）のデータの癖への対応**: 2558・2559 は 2026-06-05 に10分割されたが、日足は分割が調整されておらず、6/8 だけ更に10分の1の異常値の行がある。`Export-CalendarData.ps1` の `Repair-EtfSeries` が、「1日だけ元の水準に戻る異常値」を除き、「前日比が 1/k（k=2〜20）の跳び」を分割とみなして過去を k で割る。取得元が直したら何もしない。補正したときは実行ログに表示する。6/8 は放置パターンだけ「売買なし」の日になる。

## ファイル一覧

新規:
- `ps/Build-TwoStagePicks.ps1` … M0（`data/processed/combo/b3_b3l_m0.csv`、weight 列つき）と T2（`data/processed/combo/t2_on20.csv`）の銘柄を作る
- `ps/Patch-CalendarHold.ps1` … `web/calendar.html` の表示を放置パターンに対応させる（直してあれば何もしない。`const DATA`・`const TODAY` の行には触れない）
- `ps/config.b3l_2026_b3.json` / `config.b3l_2026_m0.json` / `config.b3l_2026_t2.json` … シミュレーションの設定

変更:
- `ps/Fetch-MarketData.ps1` … `2558.T`・`2559.T` の日足を保存する対象に追加（3行）
- `ps/Export-CalendarData.ps1` … 放置パターンを `calendar_data.json` の variants に追加（`kind = "hold"`）。`Repair-EtfSeries` を追加
- `ps/Update-B3L2026.ps1` … `Test-CrossSectionSignals.ps1` に `-AlsoShow B3` を足す／B3 の picks の変換／`Build-TwoStagePicks.ps1` の呼び出し／b3・m0・t2 のシミュレーション／`Patch-CalendarHold.ps1` の呼び出し
- `ps/Update-ForwardTest.ps1` … `$variants` に3系統を追加（各 `start = "2026-09-24"`）、読み物（`FORWARD_TEST_RESULTS.md`）に説明を追加
- `FORWARD_TEST.md` … 新系統と放置パターンの説明を追加

**触っていないもの**: `web/calendar.html` 本体（`Patch-CalendarHold.ps1` が次回の日次更新で直す。B3・M0・T2 は `Export-CalendarData.ps1` の variants に入れていないので、カレンダーには出ない）、`.github/workflows/daily-update.yml`、`Simulate-SKabu.ps1`、`Test-CrossSectionSignals.ps1`、既存の台帳の行。

エンコーディングは元のファイルにそろえてある（`.ps1` は UTF-8 BOM付き・LF、`.md` と `.json` は BOM無し・LF）。

## 適用手順

```powershell
# kabu-b2 のクローンで（パスは自分の環境に読み替える）
Copy-Item -Recurse -Force <このフォルダ>\ps\*          <クローン>\ps\
Copy-Item -Force          <このフォルダ>\FORWARD_TEST.md <クローン>\FORWARD_TEST.md
```

commit 前の確認（任意。GitHub Actions を待たずに手元で見る）:

```powershell
cd <クローン>
# 株価を展開してあること（FORWARD_TEST.md の「初回だけ」）。市場データを取り直して 2558・2559 を作る
powershell -File ps\Fetch-MarketData.ps1
# B3・M0・T2 の選定〜シミュレーション（Update-B3L2026.ps1 全体を回すなら -SkipRanking も使える）
powershell -File ps\Update-B3L2026.ps1
# 台帳への追記だけを試す（本物の台帳を書き換えたくないときは別の出力先にする）
powershell -File ps\Update-ForwardTest.ps1 -OutDir reports/_tmp_forward -SummaryMd _tmp_FORWARD.md
```

## 初回の GitHub Actions で起きること

- 新しく commit されるもの: `data/processed/b3/*.csv`、`data/processed/combo/b3_b3l_m0.csv`・`t2_on20.csv`、`data/raw/market/2558.csv`・`2559.csv`、`reports/forward_test/*`（新3系統の 2026-09-24 以降の行）、`FORWARD_TEST_RESULTS.md`、`web/calendar.html`（表示の修正＋放置パターンのデータ）、`reports/b3l_2026/calendar_data.json`。
- 新3系統の 9/24 以降の行は、パッチを反映した日の日次更新で**まとめて**足される（`recorded_at` にその日時が残る）。追記専用の台帳なので、一度書いた行は後から書き換わらない。
- 日次更新の所要時間は、シミュレーション3本と選定の追加で数分増える程度（作業コピーでの実測: シミュレーション9本で40秒、`Build-TwoStagePicks.ps1` は約7秒（既定は 2025-12-01 以降の判断日だけを作る））。

## 確認したこと（作業コピーで実行。株価は 2026-09-17 まで）

- `Test-CrossSectionSignals.ps1 -AlsoShow B3` で B3 の上位10が出る。B2 の picks は、リポジトリの既存の `daily_picks.csv` と 2023-01-05〜2026-09-15 の 905日中 784日（86.6%）で一致（残りは取得日の違いで調整後価格の履歴が違うため。私の別環境の照合では 94.8%）。
- `Build-TwoStagePicks.ps1` の T2 の上位10を、別の書き方で独立に再計算した結果と、最新の判断日（2026-09-15）で完全に一致。M0 の重み（重複銘柄が2、その他は1、19〜20銘柄）が `Simulate-SKabu.ps1` の weight 列として読まれ、平均買付銘柄数 19.5 になる。
- b3・m0・t2 のシミュレーションが通り、`Update-ForwardTest.ps1`（別の出力先・開始日を前倒しにした試験）が3系統を台帳と読み物に足す。
- `Export-CalendarData.ps1` が放置パターンを出し、`Patch-CalendarHold.ps1` を当てた `calendar.html` をブラウザで開いて、戦略ボタン・月のカレンダー・日ごとの内訳・比較表・元手/開始日の入力（開始日を変えたときの評価額は理論値と 0.03% 以内）・コンソールエラーなしを確認。
- 参考（未知データではない、2026-01-06〜09-17 のS株シミュレーション・税引後の最終資産）: B2 10銘柄 546,116円、B3 627,212円、M0 592,265円、T2 573,245円。放置は（ETF代用、税引前、〜2026-09-29）S&P500 562,049円、オルカン 560,511円。

## 確認していないこと

- GitHub Actions（Linux・PowerShell 7）での通し実行。スクリプトは PowerShell 5.1 で確認した。書き方は既存のスクリプトにそろえたが、Linux 上の実行は未確認。
- 実行時点の Yahoo が 2558・2559 を返すこと、返す日足の癖が今回確認した形（分割の未調整＋6/8 の異常値）のままであること。癖が変わっても、跳びが無ければ補正しない作りだが、別の形の異常値には対応していない。
- 東証ETF（2558・2559）と、実際の投資信託（eMAXIS Slim など）の基準価額の差（信託報酬、為替の取り方、ETFの価格の乖離）。放置パターンは「投資信託そのもの」ではなく「連動する東証ETFの分配金込み価格」で代用したもの。

## 戻したいとき

- 追加分を戻す: 適用前のファイルに戻し、`data/processed/b3/`、`data/processed/combo/b3_b3l_m0.csv`・`t2_on20.csv`、`data/raw/market/2558.csv`・`2559.csv` を削除する。台帳（`reports/forward_test/*`）に足された3系統の行は、`variant` が `b3_10`・`m0`・`t2` の行を削除して戻す（他の系統の行は触らない）。
- `web/calendar.html` は `git checkout` で戻すか、次回の日次更新で `Patch-CalendarHold.ps1` を外した状態から `Update-Calendar.ps1` を回す。
