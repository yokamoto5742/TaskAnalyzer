# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.必ず日本語で回答してください。

## 概要

WILLDOリスト(Excel)を読み込み、業務時間を集計してテンプレートExcelに出力する Windows 向け tkinter デスクトップアプリ。Python 3.13 / uv 管理。

## コマンド

```bash
uv run python main.py      # アプリ起動
uv run pyright             # 型チェック(対象は pyproject.toml の [tool.pyright])
uv run python build.py     # PyInstaller で exe をビルド(utils/config.ini を同梱)
```

テストの実行方法は `.claude/rules/testing.md` を参照。

## 構成

処理の流れ: `app/app_window.py`(GUI) → `service/service_task_analyzer.py`(全体の統括) → `service_excel_reader.py`(読み込み) → `service_data_analyzer.py`(polars で集計) → `service_excel_writer.py`(テンプレートへ書き込み)。

## 注意点

- 設定は `utils/config.ini` 一つだけ。`utils/config_manager.py` が、通常実行では自分と同じフォルダ、PyInstaller の exe では `sys._MEIPASS` からこのファイルを読む。
- テストで `load_config` をモックするときは、定義元ではなく使っている側のモジュールにパッチを当てる(例: `service.service_task_analyzer.load_config`)。
- 新しいテストは `tests/<パッケージ>/test_<モジュール>.py` に置く。`tests/` 直下にある既存のテストは、変更するときにこの配置へ移す。
