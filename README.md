# yuchodebit

ゆうちょデビット利用通知メールを Gmail から取得し、Google スプレッドシートへ明細と集計を書き出す Google Apps Script（GAS）です。

## できること

- ゆうちょデビット通知メールを検索して取得
- メール本文から以下を抽出
  - 利用日時
  - 利用店舗
  - 利用金額
- スプレッドシートに明細を書き込み
- 店舗別・日付別の合計を集計

## 前提

- Google アカウント（Gmail / スプレッドシート利用可）
- Apps Script（V8）
- 通知メール送信元: `yuchodebit@jp-bank.japanpost.jp`

## セットアップ

1. Google スプレッドシートを作成し、ID を控える
2. Apps Script プロジェクトに `script.gs` の内容を貼り付ける
3. `script.gs` の `ssId` にスプレッドシート ID を設定する
4. `newsheet` を用途に応じて設定する
   - `true`: 月初シート作成モード（その月の1日以降を対象）
   - `false`: 日次追記モード（前日以降を対象）
5. `main()` を実行して権限を承認する

## 実行

- 手動実行: `main()`
- 定期実行: Apps Script のトリガーで `main()` を設定

## 出力イメージ

- `A:D` 列: 明細（id / when / shop / expence）
- `F` 列: 総額・最大店舗・最大日付
- `H:I` 列: 日付別集計
- `J:K` 列: 店舗別集計
