import polars as pl
from datetime import datetime
from pathlib import Path
import os
from openpyxl import load_workbook
from config_manager import load_config


class TaskAnalyzer:
    def __init__(self):
        self.config = load_config()
        self.paths_config = self.config['PATHS']

    @staticmethod
    def process_excel_sheet(wb, sheet_name, date):
        """Excelシートから業務データを抽出"""
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

    def analyze_workbook(self, file_path, start_date, end_date):
        """Excelワークブックの分析"""
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
                    sheet_date = date_cell
                else:
                    sheet_date = datetime.strptime(str(date_cell), '%Y年%m月%d日')

                if start_date <= sheet_date <= end_date:  # 日付範囲のチェックを修正
                    tasks = self.process_excel_sheet(wb, sheet_name, sheet_date)
                    all_tasks.extend(tasks)
                    dates.append(sheet_date)

            except (ValueError, TypeError) as e:
                print(f"シート {sheet_name} の日付の解析でエラー: {e}")
                continue

        if not dates:
            raise ValueError("有効なデータが見つかりませんでした。")

        df = pl.DataFrame(all_tasks)
        start_date = min(dates).strftime("%Y%m%d")
        end_date = max(dates).strftime("%Y%m%d")

        return df, start_date, end_date

    def analyze_tasks(self, df, template_path, output_dir, start_date, end_date):
        """業務データの分析と結果の出力"""
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

        self._save_results(clerk_tasks, non_clerk_tasks, template_path, output_dir, start_date, end_date)

    @staticmethod
    def _save_results(clerk_tasks, non_clerk_tasks, template_path, output_dir, start_date, end_date):
        """分析結果をExcelファイルに保存"""
        wb = load_workbook(filename=template_path)

        # クラーク業務の結果を保存
        clerk_sheet = wb['クラーク業務']
        for i, row in enumerate(clerk_tasks.to_pandas().values, start=2):
            for j, value in enumerate(row, start=1):
                clerk_sheet.cell(row=i, column=j, value=value)

        # クラーク以外の業務の結果を保存
        non_clerk_sheet = wb['クラーク以外業務']
        for i, row in enumerate(non_clerk_tasks.to_pandas().values, start=2):
            for j, value in enumerate(row, start=1):
                non_clerk_sheet.cell(row=i, column=j, value=value)

        # 結果を保存して開く
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        start_date_str = datetime.strftime(datetime.strptime(str(start_date), '%Y-%m-%d %H:%M:%S'), '%Y%m%d')
        end_date_str = datetime.strftime(datetime.strptime(str(end_date), '%Y-%m-%d %H:%M:%S'), '%Y%m%d')
        output_filename = f'WILLDOリストまとめ{start_date_str}_{end_date_str}.xlsx'
        output_file_path = output_path / output_filename

        wb.save(output_file_path)
        print(f"\n集計結果を保存しました: {output_file_path}")
        os.system(f'start excel "{output_file_path}"')

    def run_analysis(self, start_date_str, end_date_str):
        """分析処理の実行"""
        try:
            start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
            end_date = datetime.strptime(end_date_str, '%Y-%m-%d')
            df, start_date_fmt, end_date_fmt = self.analyze_workbook(
                self.paths_config['input_file_path'],
                start_date,
                end_date
            )

            self.analyze_tasks(
                df,
                self.paths_config['template_path'],
                self.paths_config['output_dir'],
                start_date,
                end_date
            )

            return True
        except Exception as e:
            print(f"エラーが発生しました: {e}")
            return False
