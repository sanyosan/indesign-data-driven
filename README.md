# InDesign データドリブン スクリプト

Adobe InDesignでCSVデータを使用して、テンプレートグループを自動複製・配置・テキスト置換するスクリプトセットです。

## 概要

このスクリプトセットにより、以下の作業を自動化できます：
1. テンプレートグループの複数位置への複製
2. CSVデータに基づくテキストの自動置換
3. 位置情報の記録と再利用

## ファイル構成

### `pos.csv`
配置位置の座標データが格納されています。
- 基準グループの左上を(0,0)とした相対座標
- data1からdata8までの8つの配置位置を定義
- `get_8_positions_by_selection.jsx`で生成

### `data.csv`
テキスト置換用のデータが格納されています。
- 1行目：ヘッダー（TextFrameのlabelと対応）
- 2行目以降：各グループに適用するデータ
- 自動読み込み対象ファイル

### `advanced_data_driven.jsx`
メインの実行スクリプトです。
- pos.csvとdata.csvを自動読み込み
- グループの複製・配置・テキスト置換を実行

### `get_8_positions_by_selection.jsx`
位置データ作成用の補助スクリプトです。
- 8つのオブジェクトを順番に選択して実行
- 相対座標を計算してpos.csvを生成

### `export_group_items_to_csv.jsx`
グループ内容をCSV出力する補助スクリプトです。
- グループ内のTextFrameとRectangleの情報を抽出
- デバッグやデータ確認用

### `data_driven_copy.jsx`
旧バージョンの実行スクリプトです。
- 参考用として保持

### `test.csv`
テスト用のデータファイルです。

## 使用方法

### 初回セットアップ
1. 8つの配置位置を決めて、オブジェクトを配置
2. 8つのオブジェクトを順番に選択
3. `get_8_positions_by_selection.jsx`を実行してpos.csvを生成

### データ準備
1. `data.csv`を作成（1行目にTextFrameのlabel、2行目以降にデータ）
2. テンプレートグループ内のTextFrameにlabelを設定

### 実行
1. InDesignでテンプレートグループを選択
2. `advanced_data_driven.jsx`を実行
3. 自動的にpos.csvの位置にグループが複製され、data.csvの内容でテキストが置換される

## 必要ファイル
- `pos.csv` - 位置データ（自動読み込み）
- `data.csv` - テキストデータ（自動読み込み）
- `advanced_data_driven.jsx` - メイン実行スクリプト

## 実行方法
InDesignの「ユーティリティ」→「スクリプト」で各スクリプトを実行

## 注意事項
- TextFrameには適切なlabelを設定してください
- pos.csvとdata.csvはスクリプトと同じフォルダに配置してください
- 実行前にテンプレートグループを選択してください

## 技術仕様
- 対象：Adobe InDesign（ExtendScript）
- 座標系：geometricBounds [top, left, bottom, right]
- エンコーディング：UTF-8
- ファイル形式：CSV（カンマ区切り、ダブルクォート対応）
