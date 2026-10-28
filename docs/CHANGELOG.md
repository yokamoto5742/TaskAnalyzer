# 変更履歴

このプロジェクトのすべての重要な変更は、このファイルに記録されます。

フォーマットは [Keep a Changelog](https://keepachangelog.com/ja/1.1.0/) に基づいており、
バージョン番号は [Semantic Versioning](https://semver.org/lang/ja/) に従っています。

## [Unreleased]

## [1.2.0] - 2026-10-28
### 追加
- 「ファイルパス設定」ボタンを追加。WILLDOリスト・出力書式・出力先フォルダを画面で選んで `config.ini` に保存できる

## [1.1.3] - 2026-10-27
### 変更
- 日付を文字列に変換せず `date` のまま GUI → 読み込み → 書き込みへ渡す
- `read_workbook` の戻り値を 6 要素のタプルから `WorkbookData` データクラスにし、読み込んだ行を `TaskRecord` / `CommunicationRecord` で型付けする
- エラーの表示を GUI 側に一本化。`TaskAnalyzer.run_analysis` は保存先のパスを返し、エラーはそのまま送出する
- 期間内にデータがないときは「指定された期間内のデータがありません。」と表示する(`NoDataError`)
- エラーメッセージを `app/constants.py` に追加し、使われなくなった `UNEXPECTED_ERROR_MESSAGE` を `ANALYSIS_ERROR_MESSAGE` に置き換える
- 開始日・終了日の入力欄の生成を `_create_date_entry` にまとめ、ラベルを「開始日」「終了日」にそろえる
- `app/`・`service/`・`main.py`・`build.py` に型ヒントを追加し、import 順をルールに合わせる
- シート名などのマジック文字列をモジュール先頭の定数にする
- windowed の exe では見えない `print` によるエラー出力と、`config_manager` の `print` して再送出するだけの例外処理を削除
- `build.py` でビルドに失敗したら止まるようにする(`check=True`)

### ドキュメント
- README のファイル構成から削除した `utils/utils.py` を外し、`app/constants.py` と `utils/date_period.py` を追加

## [1.1.2] - 2026-10-27
### 変更
- 結果の書き込みで 6 シート分の同じループを 1 つにまとめる。Excel の起動を `os.system('start excel ...')` から `os.startfile` にする
- 集計式を `summarize` 関数に一本化し、`create_dataframes` と `aggregate_dataframe` を削除
- WILLDO リストの読み込みで行範囲ごとのループと時間の変換処理を共通化

### 削除
- どこからも使われていない `utils/utils.py` を削除

### テスト
- `tests/test_excel_result_writer.py` を `tests/service/test_service_excel_writer.py` へ移動

## [1.1.1] - 2026-10-27
### 修正
- 起動後に「設定ファイル」で編集した内容が、「分析開始」で起動時の内容に上書きされていた問題を修正。分析のたびに設定を読み直す
- 期間内にコミュニケーション・デイリータスクの記録が 1 件もないと、集計でエラーになっていた問題を修正
- 期間内にデータがないとき「日付の形式が正しくありません」と表示されていた問題を修正
- 全項目に含まれるコミュニケーション行で、内容と名前の間に空白がないと `内容(名前)` のまま集計されていた問題を修正
- 合計時間(`total_hours`)が切り捨てられていた(90 分 → 1 時間)問題を修正。小数第 1 位に丸める(90 分 → 1.5 時間)

### テスト
- 変更したテストを `tests/service/` へ移動

## [1.1.0] - 2026-10-27
### 追加
- 年と期間(1月〜12月・第1〜第4四半期・上半期・下半期・一年)を選ぶと、開始日と終了日が自動で入るようにする
- 年の始まりを「1月始まり」と「4月始まり」から選べるようにし、選んだ設定を `config.ini` に保存する
- 起動時に前回分析した期間を開始日・終了日に入れる

### 変更
- 開始日・終了日に `yyyy/mm/dd` 形式で直接入力した値を、フォーカスを外さなくても分析に使うようにする。形式が違うときはエラーを表示する
- エラーメッセージを `app/constants.py` にまとめる
- 期間の選択欄が収まるよう、ウィンドウの高さを 380 にする

## [1.0.4] - 2026-10-27
### 変更
- 「設定ファイル」ボタンで開くファイルを、`config.ini` の `config_path` ではなく実際に読み込んでいる設定ファイル(`CONFIG_PATH`)にする
- `config.ini` から `config_path` を削除

## [1.0.3] - 2026-10-26
### 修正
- pyright の型エラー(23件)を解消
  - 未使用の import・変数を削除
  - `sys._MEIPASS` を `getattr` で参照するよう変更
  - テストで Optional になる `wb.active` を `assert` で絞り込み

## [1.0.2] - 2026-10-26
### 削除
- どこからも使われておらず、import するとエラーになる `utils/env_loader.py` と `utils/log_rotation.py` を削除

### 修正
- pytest の `integration` マーカーを `pyproject.toml` に登録し、テスト実行時の警告を解消

### ドキュメント
- README のインストール手順を `uv sync` に変更し、ファイル構成を現在のパッケージ構成に合わせて修正

## [1.0.1] - 2026-10-26
### 修正
- パッケージ構成(`app/`・`service/`・`utils/`)への移行で壊れていた import を修正し、アプリが起動しない問題を解消
- 削除された `version.py` の代わりに `app.__version__` を参照するよう変更
- ビルド時の `config.ini` の参照先を `utils/config.ini` に修正
- テストの import とモック対象をパッケージ構成に合わせて修正

## [1.0.0] - 2026-10-26
- 初版リリース
