# コードレビュー (2026-10-27 / v1.1.0)

対象: `main.py` `build.py` `app/` `service/` `utils/`(`scripts/` と `tests/` は除外)
観点: 可読性・保守性・KISS

## サマリー

| 優先度 | 件数 | 内容 |
|---|---|---|
| 高(不具合) | 3 | 設定ファイルの上書き、空データでのクラッシュ、エラーメッセージの取り違え |
| 中(保守性) | 7 | 重複コード、日付の往復変換、使われていないモジュール、6要素タプル など |
| 低(体裁) | 6 | import順、型ヒント、print、build.py など |

全体として構成(GUI → 統括 → 読込 → 集計 → 書込)は分かりやすく、`utils/date_period.py` は小さく読みやすい良い例です。
一方 `service/` には同じ処理のコピーが多く、**約 120 行は削れます**。

---

## 高: 不具合

### H1. 「設定ファイル」で編集した内容が「分析開始」で上書きされる

`app/app_window.py:19` で起動時に読んだ `self.config` を、`start_analysis` (`:128`) でファイル全体ごと `save_config` しています。
また `TaskAnalyzer`(`service/service_task_analyzer.py:10`)も起動時に一度だけ設定を読みます。

再現手順:
1. アプリ起動 → 「設定ファイル」でメモ帳を開き、`input_file_path` を変更して保存
2. 「分析開始」を押す
3. → 古いパスで分析され、さらに `config.ini` が起動時の内容に戻る(編集が消える)

修正案: 保存直前に読み直し、日付の項目だけを書き換える。`TaskAnalyzer` も実行のたびに読み込む。

```diff
 # app/app_window.py
-            # 設定の保存
-            if 'Analysis' not in self.config:
-                self.config.add_section('Analysis')
-
-            self.config['Analysis'].update({
+            # メモ帳での編集を消さないよう、保存直前に読み直す
+            self.config = load_config()
+            self.config['Analysis'].update({
```

```diff
 # service/service_task_analyzer.py
 class TaskAnalyzer:
-    def __init__(self):
-        self.config = load_config()
-        self.paths_config = self.config['PATHS']
-        self.reader = ExcelTaskReader(self.config)
-        ...
-    def run_analysis(self, start_date_str, end_date_str):
+    def run_analysis(self, start_date: date, end_date: date) -> tuple[bool, str]:
+        config = load_config()
+        paths = config['PATHS']
+        reader = ExcelTaskReader(config)
```

### H2. 期間内にコミュニケーションの記録が 1 件もないと集計で落ちる

`pl.DataFrame([])` は列を持たないため、`group_by(['content', 'name'])` で例外になります。実際に確認しました。

```
>>> TaskDataAnalyzer().analyze_task_data(tasks, tasks, [], tasks)
ColumnNotFoundError unable to find column "name"; valid columns: []
```

クラーク業務・デイリータスクが 0 件の期間でも同じです。スキーマを指定して空でも列を持たせます。

```diff
 # service/service_data_analyzer.py
+TASK_SCHEMA = {'date': pl.Datetime, 'content': pl.String, 'minutes': pl.Float64}
+COMMUNICATION_SCHEMA = {**TASK_SCHEMA, 'name': pl.String}
 ...
-        df = pl.DataFrame(tasks)
+        df = pl.DataFrame(tasks, schema=TASK_SCHEMA)
-        comm_df = pl.DataFrame(communication_tasks)
+        comm_df = pl.DataFrame(communication_tasks, schema=COMMUNICATION_SCHEMA)
```

### H3. 「データがありません」が「日付の形式が正しくありません」と表示される

`service_excel_reader.py:142` は `ValueError("指定された期間内のデータがありません")` を投げますが、
`service_task_analyzer.py:44` が `ValueError` をすべて「日付の形式が正しくありません: …」として返します。
日付の検証は GUI 側(`parse_date`)で済んでいるので、この分岐は不要です。

```diff
-        except ValueError as ve:
-            return False, f"日付の形式が正しくありません: {str(ve)}"
         except Exception as e:
             return False, f"分析中にエラーが発生しました: {str(e)}"
```

あわせて、これらのメッセージも `app/constants.py` に移すとルール(UI メッセージは定数で一元管理)に沿います。

---

## 中: 保守性

### M1. `ExcelResultWriter.save_results` の同じループが 6 回

シート名と DataFrame の対応表にして 1 つのループにまとめます(56 行 → 約 30 行)。
`os.system('start excel ...')` は Windows 標準の `os.startfile` で十分です(シェル経由の文字列組み立てが不要)。

```diff
 class ExcelResultWriter:
+    SHEET_NAMES = ('クラーク業務', 'クラーク以外業務', 'デイリータスク',
+                   'コミュニケーション', 'コミュニケーション内容', '全項目')
+
     @staticmethod
     def save_results(analysis_results, template_path, output_dir, start_date, end_date):
-        clerk_tasks, non_clerk_tasks, daily_tasks, communication_by_name, \
-        communication_by_content, all_items_summary = analysis_results
-
         wb = load_workbook(filename=template_path)
-
-        clerk_sheet = wb['クラーク業務']
-        for i, row in enumerate(clerk_tasks.iter_rows(), start=2):
-            ...(同様のブロックが 6 つ)
+        for sheet_name, result in zip(ExcelResultWriter.SHEET_NAMES, analysis_results):
+            sheet = wb[sheet_name]
+            for row in result.iter_rows():
+                sheet.append(row)  # ※テンプレートが見出し行のみなら 2 行目から追記される
 ...
-        os.system(f'start excel "{output_file_path}"')
+        os.startfile(output_file_path)
```

> `sheet.append` はシートの最終行の次に書き込みます。テンプレートに見出し以外の書式行がある場合は、元の `cell(row=i, column=j)` 方式を 1 つのヘルパー関数にまとめる形にしてください。

### M2. 集計式 `agg([...])` が 3 か所にコピーされている

`service_data_analyzer.py:23-27, 33-37, 68-72` は同一です。また `aggregate_dataframe` は
「`group_by_col == 'content'` かつ `name` 列がある」という暗黙の条件で分岐しており、読み手が挙動を追いにくくなっています。
グループ化する列を明示的に渡す形にすると、分岐も `communication_by_content` の特別扱いも消えます。

```python
SUMMARY_COLUMNS = [
    pl.col('minutes').sum().alias('total_minutes'),
    (pl.col('minutes').sum() / 60).cast(pl.Int64).alias('total_hours'),
    pl.col('minutes').count().alias('frequency'),
]


def summarize(df: pl.DataFrame, keys: list[str]) -> pl.DataFrame:
    """keys ごとに時間と回数を集計し、keys の先頭列 → 合計時間の降順で並べる"""
    return (
        df.group_by(keys)
        .agg(SUMMARY_COLUMNS)
        .sort([keys[0], 'total_minutes'], descending=[False, True])
    )
```

呼び出し側:

```python
is_clerk = pl.col('content').str.contains('クラーク業務')
clerk_tasks = summarize(df.filter(is_clerk), ['content'])
non_clerk_tasks = summarize(df.filter(~is_clerk), ['content'])
daily_tasks_agg = summarize(daily_df, ['content'])
communication_by_name = summarize(comm_df, ['name'])
communication_by_content = summarize(comm_df, ['name', 'content']).select(
    ['content', 'name', 'total_minutes', 'total_hours', 'frequency'])
all_items_summary = summarize(all_items_df, ['content'])
```

> 1 キーの場合は現状「合計時間の降順」のみ、上記は「キー昇順 → 合計時間降順」になり並び順が変わります。
> 1 キーのときは `sort('total_minutes', descending=True)` を残すなど、出力 Excel の並びを変えたくない場合は調整してください。

`create_dataframes` は `pl.DataFrame(...)` を 4 回呼ぶだけなので、`analyze_task_data` 内に直接書けば十分です。

※ `total_hours` は `cast(pl.Int64)` で**切り捨て**(90 分 → 1 時間)です。意図どおりか確認してください。

### M3. `ExcelTaskReader` の行範囲ループが 4 回

`load_excel_task_data` と `load_excel_sheet_all_items` で「設定から開始・終了行を読む → ループ → `extract_cell_data`」が繰り返されています。
行範囲を受け取るヘルパーに寄せます。

```python
def _read_rows(self, sheet, start_key: str, end_key: str, sheet_date: datetime) -> list[dict]:
    start_row = self.config.getint('Analysis', start_key)
    end_row = self.config.getint('Analysis', end_key)
    rows = (self.extract_cell_data(sheet, row, sheet_date) for row in range(start_row, end_row + 1))
    return [row for row in rows if row]
```

また `extract_cell_data` の時間の変換は二重の `try` になっていますが、次で同じ意味です。

```python
try:
    minutes = float(time)
except (ValueError, TypeError):
    return None
```

`content.split()[0] if content else content` の `if content` は直前で確認済みなので不要です。

### M4. 読込 → 統括で日付を「datetime → 文字列 → datetime」と往復している

- GUI: `date` → `strftime('%Y-%m-%d')` → `TaskAnalyzer`: `strptime`
- Reader: `min(dates).strftime('%Y%m%d')` → `TaskAnalyzer`: `strptime('%Y%m%d')` → Writer: `strftime('%Y%m%d')`

`date` / `datetime` のまま受け渡せば、`strptime` 4 か所と書式文字列の重複が消えます。
(Reader 内の比較は `sheet_date.date()` と `date` で比較すれば型がそろいます。)

### M5. `read_workbook` が 6 要素のタプルを返す

呼び出し側で順番を覚える必要があり、取り違えても型チェックで検出できません。`dataclass` にまとめると読みやすくなります。

```python
@dataclass
class WorkbookData:
    tasks: list[dict]
    daily_tasks: list[dict]
    communication_tasks: list[dict]
    all_items: list[dict]
    first_date: datetime
    last_date: datetime
```

集計結果(6 つの DataFrame)も同様ですが、M1 のように `zip` でシート名と対応付けるならタプルのままでも構いません。

### M6. `utils/utils.py` はどこからも使われていない

`parse_date_safely` `safe_float_conversion` `extract_name_from_content` はいずれも参照ゼロで、
`service_excel_reader.py` 内に同じ処理のコピーがあります。どちらかに一本化してください
(リーダー内に置き、`utils/utils.py` は削除するのが最もシンプルです)。

### M7. 例外処理が二重・過剰

- `app_window.start_analysis` の外側の `try/except Exception` と、`TaskAnalyzer.run_analysis` の `except Exception` が重なっています。どちらか一方(GUI 側)で十分です。
- `config_manager.load_config` / `save_config` は例外ごとに `print` して再送出しているだけです。`--windowed` の exe では `print` は見えないため、`try` ごと削除して呼び出し側のメッセージ表示に任せられます。

```diff
 def load_config() -> configparser.ConfigParser:
     config = configparser.ConfigParser()
-    try:
-        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
-            config.read_file(f)
-    except FileNotFoundError:
-        print(...)
-        raise
-    ...
+    with open(CONFIG_PATH, encoding='utf-8') as f:
+        config.read_file(f)
     return config
```

---

## 低: 体裁・細部

| # | 場所 | 指摘 |
|---|---|---|
| L1 | `app/app_window.py:1-12` | import 順がルール違反(`subprocess` が `tkcalendar` の後、ローカルがアルファベット順でない) |
| L2 | `app_window.py:43-54` | 開始日・終了日の `DateEntry` 生成が同じ記述。`_create_date_entry(parent, row, label, initial)` にまとめられる。ラベルも「開始日」「終了日:」でコロンの有無が不統一 |
| L3 | 全体 | 型ヒント必須のルールに対し、`app_window.py` の多くのメソッド、`service/` 全体、`main.py` に型ヒントがない |
| L4 | `service_excel_reader.py` | `print` によるエラー出力は windowed exe では見えない。読み飛ばしたシートを知りたいなら `logging` でファイルに出すか、不要なら削除 |
| L5 | `service_excel_reader.py:115` | `'シート一覧'` などシート名・`'クラーク業務'` がマジック文字列。モジュール先頭の定数にする |
| L6 | `build.py` | `subprocess.run` に `check=True` がなく、失敗しても「successfully」と表示される。`f"..."` にプレースホルダーがない |

---

## 確認したい点(仕様)

1. **数式セル**: `load_workbook` に `data_only=True` がないため、C 列が数式(`=SUM(...)` 等)だと値ではなく数式文字列が返り、`float` 変換に失敗して**黙って読み飛ばされます**。WILLDO リストで C 列に数式を使っていないか確認してください。使っていなければ `read_only=True, data_only=True` を付けると読み込みも速くなります。
2. **全項目の範囲**: `load_excel_sheet_all_items` は `start_row`〜`daily_task_end_row` を読むため、コミュニケーション行(26〜34)も含まれます。この行は `内容(名前)` 形式なので、空白なしの場合 `content` に `(名前)` が残ったまま集計されます。意図どおりでしょうか。
3. **`total_hours` の切り捨て**(M2 参照)。

---

## 推奨する進め方

1. H1〜H3 を修正(再現テストを先に追加)
2. M6(未使用モジュール削除)→ M1 → M2 → M3 の順に重複を削減。各ステップで `uv run pytest tests/ -v --tb=short` が通ることを確認
3. M4・M5 で日付と戻り値の型を整理し、L3 の型ヒントを追加して `uv run pyright` を通す
4. 触ったテストは `tests/<パッケージ>/test_<モジュール>.py` へ移動(CLAUDE.md の方針)
