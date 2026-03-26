# Boostキャンペーン停止オペレーション自動化（GAS）

## 概要
このGASは、Amazon SPのBoostキャンペーンを停止する際に、好調キーワードだけを通常のManualキャンペーンへ完全一致で移行し、停止用の一括アップロードファイルを作るためのツールです。

できること:
- 停止対象ASINに紐づくBoostキャンペーンを `Campaign` 単位で `archived` 出力
- 好調な `keyword` を通常の `Manual` キャンペーンへ `exact` で移行
- 移行先候補が曖昧な場合は `output_blockers` に候補を表示
- `output_bulk_sp` からAmazonアップロード用ファイルを作成

## 前提
- クライアント側で、移行先の通常キャンペーンと広告グループはすでに存在している
- Boost停止対象のASINは、担当者が `input_stop_products` で指定する
- キーワード移行先は `Manual` キャンペーンのみ
- `Auto` キャンペーンには移行しない

## 週次の標準手順
1. Amazon広告のバルクをダウンロードする
2. `スポンサープロダクト広告キャンペーン` シートを使う
3. できれば次の行を残してCSV化する
- `キーワード`
- `商品ターゲティング`
- `商品広告`  
`商品広告` 行はASIN補完のためにあると安全です。

4. GASスプレッドシートの `input_bulk_sp_raw` にインポートする
5. `input_stop_products` に停止するASINを入れ、`approved_to_stop=true`
6. `Boost Ops` → `1) 標準実行（RAW取込後）`
7. 必要なら `output_blockers` の `selected_candidate` で移行先候補を選ぶ
8. 候補を選んだ場合は `Boost Ops` → `2) blocker反映して一括行を再生成`
9. `output_bulk_sp` を確認する
10. `Boost Ops` → `3) アップロード用Excelを作成`
11. 作成された `xlsx` をAmazonへアップロードする

## 最初に1回だけやること
1. `Boost Ops` → `0) 初期セットアップ`
2. 必要に応じて `config` を調整する

## 複数担当者への配布方法
おすすめは、完成版スプレッドシートをテンプレートとして残し、担当者ごと・案件ごとにコピーして使う方法です。

配布手順:
1. 管理者が完成版スプレッドシートをテンプレートとして保管する
2. 担当者はテンプレートを直接編集せず、`ファイル → コピーを作成` で自分用シートを作る
3. コピーしたシートで `Boost Ops` メニューが表示されることを確認する
4. 最初に `Boost Ops` → `0) 初期セットアップ` を実行する
5. 案件ごとの `config`、`input_stop_products`、`input_bulk_sp_raw` を入力して運用する

運用ルール:
- 1案件につき1スプレッドシートを使う
- テンプレート本体は管理者のみ編集する
- 担当者は必ずコピー側で作業する
- `input_stop_products` は案件ごとに手入力で管理する

担当者への案内文例:
- このシートはテンプレートです。直接編集せず、必ずコピーを作成してから利用してください。
- 最初に `Boost Ops` → `0) 初期セットアップ` を実行してください。
- 案件ごとに別シートで管理してください。

## Amazonバルクの取得条件
- 期間: `60日`
- フィルター: `Sponsored Products: Targeting and Keyword Filter = Enabled`
- 主に使うのは `スポンサープロダクト広告のデータ`
- `商品広告` 行も含めるとASIN補完が安定する

## 主要シート
- `config`: 判定条件と入札係数
- `input_bulk_sp_raw`: Amazonバルク生データ
- `input_stop_products`: 停止承認ASIN
- `input_mapping`: 移行先マッピング
- `input_boost`: Boost実績データ
- `output_migrate_exact`: 移行するキーワード
- `output_pause_boost`: アーカイブするBoostキャンペーン
- `output_blockers`: 自動判定できなかった移行候補
- `output_bulk_sp`: 最終アップロード用データ

## 重要な仕様
### 停止
- Boost停止は常に `Campaign` 単位
- 出力は `Campaign / Update / archived`
- `pause_level` は互換用の設定値として残っているが、停止出力は常にCampaign

### 移行
- 移行対象は `keyword` のみ
- `product_targeting` は移行しない
- 移行先は `Manual` キャンペーンのみ
- `Auto` / `オート` / `boost` / `全商品` / `B2B` を含むキャンペーンは移行先から除外
- 基本は「同じ広告グループ名」の通常キャンペーンへ移行
- バリエーションでASINがずれていても、同じ広告グループ名のManual候補が1件なら自動採用

### blocker対応
- `output_blockers` に候補が出る場合がある
- `selected_candidate` のプルダウンで候補を選べる
- 選択後に `2) blocker反映して一括行を再生成` を実行すると、その選択分も `output_bulk_sp` に入る

## configでよく使う項目
- `min_clicks_to_migrate`
- `min_orders_to_migrate`
- `min_roas_to_migrate`
- `min_cvr_to_migrate`
- `fallback_bid`
- `use_roas_bid_multiplier`
- `roas_bid_multiplier_ge_35`
- `roas_bid_multiplier_ge_25`
- `roas_bid_multiplier_ge_20`
- `roas_bid_multiplier_ge_15`
- `roas_bid_multiplier_lt_15`
- `min_bid`
- `max_bid`

ROAS目標20の標準値:
- `use_roas_bid_multiplier = true`
- `roas_bid_multiplier_ge_35 = 0.85`
- `roas_bid_multiplier_ge_25 = 0.75`
- `roas_bid_multiplier_ge_20 = 0.65`
- `roas_bid_multiplier_ge_15 = 0.55`
- `roas_bid_multiplier_lt_15 = 0.45`
- `min_bid = 20`
- `max_bid = 120`

## 出力前の最終確認
- `output_bulk_sp` にキーワード移行先として `Auto` が入っていない
- 移行行が `Keyword / Create / exact`
- 停止行が `Campaign / Update / archived`
- 入札額が `0` になっていない
- `output_blockers` が空、または内容を確認済み

## メニュー一覧
- `0) 初期セットアップ`
- `1) 標準実行（RAW取込後）`
- `2) blocker反映して一括行を再生成`
- `3) アップロード用Excelを作成`

## 補足
- `1)` は `RAWから必要データ作成` と `判定＆出力作成` を続けて実行する
- `3)` はAmazonアップロード用に `output_bulk_sp` だけを `xlsx` 出力する
- `input_stop_products` は自動更新しない。停止するASINは毎回手入力で管理する
- `input_mapping` は `1)` の中でRAWから自動更新されるが、必要に応じて担当者が微修正してよい
