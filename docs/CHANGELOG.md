# 変更履歴

このプロジェクトのすべての重要な変更は、このファイルに記録されます。

フォーマットは [Keep a Changelog](https://keepachangelog.com/ja/1.1.0/) に基づいており、
バージョン番号は [Semantic Versioning](https://semver.org/lang/ja/) に従っています。

## [Unreleased]

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
