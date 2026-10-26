# 変更履歴

このプロジェクトのすべての重要な変更は、このファイルに記録されます。

フォーマットは [Keep a Changelog](https://keepachangelog.com/ja/1.1.0/) に基づいており、
バージョン番号は [Semantic Versioning](https://semver.org/lang/ja/) に従っています。

## [Unreleased]

## [1.0.1] - 2026-10-26
### 修正
- パッケージ構成(`app/`・`service/`・`utils/`)への移行で壊れていた import を修正し、アプリが起動しない問題を解消
- 削除された `version.py` の代わりに `app.__version__` を参照するよう変更
- ビルド時の `config.ini` の参照先を `utils/config.ini` に修正
- テストの import とモック対象をパッケージ構成に合わせて修正

## [1.0.0] - 2026-10-26
- 初版リリース
