# Boostキャンペーン停止オペレーション自動化（GAS）

## できること
- Boost実績をもとに、停止対象（CampaignまたはAd Group）を判定
- 停止前に、好調ターゲットを通常キャンペーンへ`完全一致`で移行する行を作成
- 移行先が無い好調ターゲットは`blocker`として出力し、停止を自動でブロック
- クライアント判断で停止OKになった商品（`product_key`）だけを処理対象にできる
- `output_migrate_exact` と `output_pause_boost` から、SP向けの統合アップロード行 `output_bulk_sp` を自動生成
- 出力シートをCSV化してDriveへ保存

## 前提
- クライアント側で新商品のBoost/通常キャンペーン作成と商品追加は実施済み
- このGASは「停止判定」「移行行作成」「停止行作成」に特化
- Amazon Adsの実アップロード列名はアカウントのテンプレに合わせる必要あり
- 本ツール出力をそのまま使うか、既存Bulkテンプレへマッピングして利用

## データダウンロード条件（固定運用）
Amazon広告の「バルク操作」画面で、以下の条件でダウンロードする。

- パフォーマンス指標の期間: `60日`
- `Sponsored Products: Targeting and Keyword Filter`: `Enabled`
- 含む（チェックON）:
- `スポンサープロダクト広告のデータ`
- `スポンサーブランド広告のデータ`
- `スポンサーブランド広告の複数広告グループのデータ`
- `スポンサーディスプレイ広告のデータ`
- 含む（チェックOFF）:
- `終了したキャンペーン`
- `Paused campaigns`
- `インプレッションがないキャンペーンアイテム`
- `キャンペーンの掲載枠のデータ`
- `ブランドアセットデータ`
- `スポンサープロダクト広告に関するガイダンス`
- `language`
- `スポンサープロダクト広告の検索用語データ`
- `スポンサーブランド広告の検索用語データ`
- `Budget Rules Data`

補足:
- このGASで主に使うのはSP（スポンサープロダクト）データ。
- 立ち上がりが遅い商品を拾うため、期間は60日を標準とする。

## 導入
1. Googleスプレッドシートを作成
2. 拡張機能 → Apps Script を開く
3. `Code.gs` を貼り付けて保存
4. シートに戻って再読み込み
5. メニュー `Boost Ops` → `0) シート初期化`

## シート構成
- `config`: 判定しきい値
- `input_bulk_sp_raw`: Amazonバルク（SPシート）生データ貼り付け
- `input_boost_campaigns`: Boostキャンペーン判定マスタ（ID/名称キーワード/商品キー）
- `input_boost`: Boost実績
- `input_mapping`: 移行先マスタ
- `input_stop_products`: 停止承認済み商品リスト
- `output_migrate_exact`: 通常キャンペーンへの追加行
- `output_pause_boost`: Boost停止行
- `output_blockers`: 停止をブロックした理由
- `output_summary`: 件数サマリ
- `output_bulk_sp`: Amazonテンプレート「スポンサープロダクト広告キャンペーン」シートの列順に合わせた統合一括アップロード行

## 入力: input_bulk_sp_raw
- Amazonバルクの「スポンサープロダクト広告キャンペーン」シートをそのまま貼り付け
- メニュー `0.5) SP生データをinput_boostへ変換` を実行すると `input_boost` を自動作成

## 入力: input_boost_campaigns
- `boost_campaign_id`（推奨）
- `boost_campaign_name_contains`（任意。名称部分一致）
- `product_key`（ASIN/型番）
- `product_label`（任意）
- `enabled`（空欄または true で有効）

補足:
- まずは `boost_campaign_id` で管理するのが最も安全。
- ID未設定でも `boost_campaign_name_contains` と `config.boost_campaign_name_contains` でBoost判定可能。

## 入力: input_boost
必須列:
- `boost_campaign_id`
- `boost_campaign_name`
- `boost_ad_group_id`
- `boost_ad_group_name`
- `product_key`（ASIN/型番など。停止承認制御に使用）
- `target_type` (`keyword` or `product_targeting`)
- `target_text`
- `impressions`
- `clicks`
- `orders`
- `sales`
- `spend`

任意列:
- `product_label`（商品名など任意ラベル）
- `match_type`
- `cpc`
- `target_key`（明示キー。未設定時は `target_type::target_text` を内部生成）
- `existing_exact_in_normal`（`true/1/yes` なら移行済み扱い）

作成ルール:
- `product_key` はASINまたは型番で統一する（週次で同じキー体系を維持）。
- 1行は「Boost内のターゲット実績」単位で入れる。
- 文字列前後の空白は削除してから貼り付ける。

## 入力: input_mapping
- `target_key`（推奨）
- `target_type`
- `target_text`
- `normal_campaign_id`
- `normal_campaign_name`
- `normal_ad_group_id`
- `normal_ad_group_name`
- `default_bid`
- `state`（空なら `enabled`）

`target_key` が空なら `target_type::target_text` で照合します。

## 入力: input_stop_products
- `product_key`（`input_boost.product_key` と一致）
- `product_label`（任意）
- `approved_to_stop`（`true/1/yes` で停止承認）
- `note`（任意）

## 判定ロジック
### 移行判定（好調ターゲット）
以下をすべて満たすと、通常キャンペーンへ`exact`で移行:
- `clicks >= min_clicks_to_migrate`
- `orders >= min_orders_to_migrate`
- `ACOS <= max_acos_to_migrate`
- `CVR >= min_cvr_to_migrate`

補足:
- `ACOS = spend / sales * 100`
- `CVR = orders / clicks * 100`

### 停止判定（Boost）
- 停止単位（`pause_level`）で`impressions`合計が`min_impressions_to_stop`以上
- かつ、停止単位内で「移行必須なのに移行先が無いターゲット」が0件

## 実行手順（週次）
1. `input_bulk_sp_raw` にSP生データを貼る（Amazonバルク2枚目）
2. `input_boost_campaigns` を更新（BoostキャンペーンIDと product_key を紐付け）
3. `Boost Ops` → `0.5) SP生データをinput_boostへ変換`
4. `input_mapping` を更新
5. `input_stop_products` にクライアント承認済みの停止対象商品を貼る（`approved_to_stop=true`）
6. `Boost Ops` → `1) 判定＆出力作成`
7. `output_blockers` を確認し、必要ならマッピング追加または `product_key` を補完
8. `Boost Ops` → `1) 判定＆出力作成` を再実行
9. `output_bulk_sp` をBulkアップロード（または `output_migrate_exact` と `output_pause_boost` を個別利用）
10. 必要なら `2) 出力シートをCSV化（Drive保存）`

補助:
- `1.5) Amazon SP一括行を再生成` は、判定結果を変えずに `output_bulk_sp` だけ再作成したい場合に利用

## config 初期値
- `min_impressions_to_stop`: 3000
- `min_clicks_to_migrate`: 8
- `min_orders_to_migrate`: 1
- `max_acos_to_migrate`: 35
- `min_cvr_to_migrate`: 8
- `fallback_bid`: 45
- `bulk_product_type`: `Sponsored Products`
- `boost_campaign_name_contains`: `boost,ブースト`（生データ変換時の補助判定）
- `require_stop_approval`: `true`（停止承認済み商品だけ処理）
- `pause_level`: `ad_group`（`campaign`も可）

## 運用上の注意
- 新規SKUが多い週は、まず`output_blockers`を0にしてから停止実行
- `existing_exact_in_normal`を正しく入れると重複移行を抑制
- `output_bulk_sp` はSP一括アップロードの共通列に寄せた形式です。アカウントテンプレートで列名・必須列が異なる場合は調整してください
- 今回の実装は、共有いただいたテンプレート（2枚目シート）のヘッダーに合わせています
- 本番前に少量データで試験し、貴社のBulkテンプレ列に合わせて最終調整を推奨

## アシスタント運用チェックリスト
- 60日・Filter Enabled のダウンロード条件になっているか確認
- `input_boost_campaigns` に Boost campaign ID と `product_key` を紐付け済みか確認
- `0.5) SP生データをinput_boostへ変換` を先に実行
- `input_boost` に `product_key` が欠損していないか確認
- `input_stop_products` の承認フラグが `true/1/yes` か確認
- `output_blockers` が0件になるまで再実行
- 最終的に `output_bulk_sp` をアップロード
