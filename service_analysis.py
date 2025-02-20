import polars as pl
from datetime import datetime
from pathlib import Path
import os
import math
from openpyxl import load_workbook
from config_manager import load_config


def process_excel_sheet(wb, sheet_name, date):
    sheet = wb[sheet_name]

    tasks = []
    for row in range(4, 24):
        content = sheet[f'B{row}'].value
        time = sheet[f'C{row}'].value

        if content and time and time != '*':
            try:
                content = content.split()[0] if content else content
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
    all_tasks = []
    dates = []

    for sheet_name in wb.sheetnames:
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
                dates.append(date)

        except (ValueError, TypeError) as e:
            print(f"シート {sheet_name} の日付の解析でエラー: {e}")
            continue

    df = pl.DataFrame(all_tasks)

    if not dates:
        raise ValueError("有効なデータが見つかりませんでした。")

    start_date = min(dates).strftime("%Y%m%d")
    end_date = max(dates).strftime("%Y%m%d")

    return df, start_date, end_date


def analyze_tasks(df, template_path, output_dir, start_date, end_date):
    # クラーク業務の集計
    clerk_tasks = (
        df.filter(pl.col('content').str.contains('クラーク業務'))
        .group_by('content')
        .agg([
            pl.col('minutes').sum().alias('total_minutes'),
            (pl.col('minutes').sum() / 60).cast(pl.Int64).alias('total_hours'),
            pl.col('minutes').count().alias('frequency')
        ])
        .sort('total_minutes', descending=True)
    )

    # クラーク以外の業務の集計
    non_clerk_tasks = (
        df.filter(~pl.col('content').str.contains('クラーク業務'))
        .group_by('content')
        .agg([
            pl.col('minutes').sum().alias('total_minutes'),
            (pl.col('minutes').sum() / 60).cast(pl.Int64).alias('total_hours'),
            pl.col('minutes').count().alias('frequency')
        ])
        .sort('total_minutes', descending=True)
    )

    wb = load_workbook(filename=template_path)

    clerk_sheet = wb['クラーク業務']
    for i, row in enumerate(clerk_tasks.to_pandas().values, start=2):
        for j, value in enumerate(row, start=1):
            clerk_sheet.cell(row=i, column=j, value=value)

    non_clerk_sheet = wb['クラーク以外業務']
    for i, row in enumerate(non_clerk_tasks.to_pandas().values, start=2):
        for j, value in enumerate(row, start=1):
            non_clerk_sheet.cell(row=i, column=j, value=value)

    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    output_filename = f'WILLDOリストまとめ{start_date}_{end_date}.xlsx'
    output_file_path = output_path / output_filename

    wb.save(output_file_path)
    print(f"\n集計結果を保存しました: {output_file_path}")

    os.system(f'start excel "{output_file_path}"')


def main():
    try:
        config = load_config()
        paths_config = config['PATHS']

        df, start_date, end_date = analyze_workbook(paths_config['input_file_path'])

        analyze_tasks(
            df,
            paths_config['template_path'],
            paths_config['output_dir'],
            start_date,
            end_date
        )

    except Exception as e:
        print(f"エラーが発生しました: {e}")


if __name__ == "__main__":
    main()
