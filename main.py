import polars as pl
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from pathlib import Path


def process_excel_sheet(wb, sheet_name, date):
    sheet = wb[sheet_name]

    # B3:C23の範囲からデータを取得
    tasks = []
    for row in range(4, 24):
        content = sheet[f'B{row}'].value
        time = sheet[f'C{row}'].value

        if content and time and time != '*':  # 両方のセルに値があり、時間が'*'でない場合のみ処理
            try:
                minutes = float(time)
                tasks.append({
                    'date': date,
                    'content': content,
                    'minutes': minutes
                })
            except (ValueError, TypeError) as e:
                print(f"時間の変換でエラー: {e}")

    return tasks


def analyze_workbook(file_path):
    wb = load_workbook(filename=file_path)

    # 全シートのデータを収集
    all_tasks = []
    for sheet_name in wb.sheetnames:
        # シート一覧は処理をスキップ
        if sheet_name == 'シート一覧':
            continue

        sheet = wb[sheet_name]
        date_cell = sheet['A1'].value

        try:
            if isinstance(date_cell, datetime):
                date = date_cell
            else:
                date = datetime.strptime(str(date_cell), '%Y年%m月%d日')

            if datetime(2025, 1, 1) <= date <= datetime(2025, 12, 31):
                tasks = process_excel_sheet(wb, sheet_name, date)
                all_tasks.extend(tasks)

        except (ValueError, TypeError) as e:
            print(f"シート {sheet_name} の日付の解析でエラー: {e}")
            continue

    # Polarsデータフレームを作成
    df = pl.DataFrame(all_tasks)

    # 業務内容ごとの集計
    summary = df.group_by('content').agg([
        pl.col('minutes').sum().alias('total_minutes'),
        pl.len().alias('frequency')
    ]).sort('total_minutes', descending=True)

    return summary


def save_to_csv(summary, output_dir):
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Polars DataFrameをpandas DataFrameに変換して保存
    summary.to_pandas().to_csv(output_dir / 'summary_task.csv', encoding='utf-8-sig', index=False)


def main():
    file_path = r"C:\Shinseikai\TaskAnalyzer\WILLDOリスト.xlsx"  # Excelファイルのパス
    output_dir = "analysis_output"  # 出力ディレクトリ
    summary= analyze_workbook(file_path)
    save_to_csv(summary, output_dir)


if __name__ == "__main__":
    main()
