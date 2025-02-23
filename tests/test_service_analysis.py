import pytest
from datetime import datetime
import polars as pl
from unittest.mock import MagicMock, patch
from openpyxl import Workbook
from pathlib import Path

from service_analysis import TaskAnalyzer


@pytest.fixture
def mock_config():
    from configparser import ConfigParser
    config = ConfigParser()

    config['PATHS'] = {
        'input_file_path': 'dummy/path/willdo.xlsx',
        'template_path': 'dummy/path/template.xlsx',
        'output_dir': 'dummy/path/output'
    }
    config['Analysis'] = {
        'start_row': '4',
        'end_row': '24',
        'communication_start_row': '26',
        'communication_end_row': '34',
        'daily_task_start_row': '37',
        'daily_task_end_row': '43'
    }
    return config


@pytest.fixture
def analyzer(mock_config):
    with patch('service_analysis.load_config', return_value=mock_config):
        return TaskAnalyzer()


@pytest.fixture
def sample_sheet():
    wb = Workbook()
    sheet = wb.active

    # テスト用のデータを設定
    sheet['A1'] = '2025年2月23日'
    sheet['B4'] = 'クラーク業務1'
    sheet['C4'] = 30
    sheet['B5'] = '一般業務1'
    sheet['C5'] = 45
    sheet['B37'] = 'デイリータスク1'
    sheet['C37'] = 15
    sheet['B26'] = '打ち合わせ(山田)'
    sheet['C26'] = 60

    return sheet


def test_extract_cell_data(analyzer, sample_sheet):
    date = datetime(2025, 2, 23)
    result = analyzer.extract_cell_data(sample_sheet, 4, date)

    assert result is not None
    assert result['content'] == 'クラーク業務1'
    assert result['minutes'] == 30
    assert result['date'] == date


def test_extract_cell_data_invalid(analyzer, sample_sheet):
    date = datetime(2025, 2, 23)
    # 空のセルの場合
    sample_sheet['B6'] = None
    sample_sheet['C6'] = None
    result = analyzer.extract_cell_data(sample_sheet, 6, date)
    assert result is None

    # 時間が'*'の場合
    sample_sheet['B7'] = '業務3'
    sample_sheet['C7'] = '*'
    result = analyzer.extract_cell_data(sample_sheet, 7, date)
    assert result is None


def test_load_excel_task_data(analyzer, sample_sheet):
    date = datetime(2025, 2, 23)
    wb = MagicMock()
    wb.__getitem__.return_value = sample_sheet

    tasks, daily_tasks, comm_tasks = analyzer.load_excel_task_data(wb, 'Sheet1', date)

    assert len(tasks) > 0
    assert len(daily_tasks) > 0
    assert len(comm_tasks) > 0

    # クラーク業務のチェック
    assert any(task['content'] == 'クラーク業務1' for task in tasks)
    # デイリータスクのチェック
    assert any(task['content'] == 'デイリータスク1' for task in daily_tasks)
    # コミュニケーションタスクのチェック
    assert any(task['name'] == '山田' for task in comm_tasks)


def test_aggregate_dataframe(analyzer):
    # テスト用のデータフレームを作成
    data = {
        'content': ['タスク1', 'タスク1', 'タスク2'],
        'minutes': [30, 45, 60],
        'date': [datetime(2025, 2, 23)] * 3
    }
    df = pl.DataFrame(data)

    result = analyzer.aggregate_dataframe(df)

    assert len(result) == 2
    assert result['total_minutes'][0] == 75  # タスク1の合計
    assert result['frequency'][0] == 2  # タスク1の出現回数


@pytest.mark.integration
def test_run_analysis(analyzer):
    with patch('service_analysis.load_workbook'), \
            patch('service_analysis.TaskAnalyzer.analyze_workbook') as mock_analyze_wb, \
            patch('service_analysis.TaskAnalyzer.analyze_tasks') as mock_analyze_tasks:
        # モックの戻り値を設定
        mock_analyze_wb.return_value = (
            pl.DataFrame({'content': [], 'minutes': [], 'date': []}),  # df
            pl.DataFrame({'content': [], 'minutes': [], 'date': []}),  # daily_df
            pl.DataFrame({'name': [], 'minutes': [], 'date': []}),  # comm_df
            pl.DataFrame({'content': [], 'minutes': [], 'date': []}),  # all_items_df
            '20250223',  # start_date_fmt
            '20250223'  # end_date_fmt
        )

        success, error = analyzer.run_analysis('2025-02-23', '2025-02-23')

        assert success is True
        assert error is None
        mock_analyze_wb.assert_called_once()
        mock_analyze_tasks.assert_called_once()


def test_run_analysis_invalid_date(analyzer):
    success, error = analyzer.run_analysis('invalid-date', '2025-02-23')
    assert success is False
    assert '日付の形式が正しくありません' in error


if __name__ == '__main__':
    pytest.main(['-v'])
