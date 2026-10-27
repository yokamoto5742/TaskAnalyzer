from datetime import date
from unittest.mock import patch

import pytest

from service.service_excel_reader import WorkbookData
from service.service_task_analyzer import TaskAnalyzer

PATHS = {'input_file_path': 'test_input.xlsx', 'template_path': 'test_template.xlsx', 'output_dir': 'test_output'}


@pytest.fixture
def mocks():
    with patch('service.service_task_analyzer.load_config') as mock_load_config, \
            patch('service.service_task_analyzer.ExcelTaskReader') as mock_reader_class, \
            patch('service.service_task_analyzer.TaskDataAnalyzer') as mock_analyzer_class, \
            patch('service.service_task_analyzer.ExcelResultWriter') as mock_writer_class:
        mock_load_config.return_value = {'PATHS': PATHS}
        yield mock_load_config, mock_reader_class.return_value, mock_analyzer_class.return_value, mock_writer_class


class TestTaskAnalyzer:
    def test_run_analysis_success(self, mocks):
        _, mock_reader, mock_analyzer, mock_writer = mocks
        workbook_data = WorkbookData(
            tasks=[{'date': date(2024, 1, 1), 'content': 'タスク1', 'minutes': 30}],
            daily_tasks=[{'date': date(2024, 1, 1), 'content': '日次タスク1', 'minutes': 10}],
            communication_tasks=[{'date': date(2024, 1, 1), 'name': '田中', 'content': '打合せ', 'minutes': 30}],
            all_items=[{'date': date(2024, 1, 1), 'content': 'すべて1', 'minutes': 30}],
            first_date=date(2024, 1, 2),
            last_date=date(2024, 1, 4),
        )
        mock_reader.read_workbook.return_value = workbook_data
        mock_analyzer.analyze_task_data.return_value = ('results',)
        mock_writer.save_results.return_value = 'output_file_path.xlsx'

        output_file = TaskAnalyzer().run_analysis(date(2024, 1, 1), date(2024, 1, 5))

        assert output_file == 'output_file_path.xlsx'
        mock_reader.read_workbook.assert_called_once_with('test_input.xlsx', date(2024, 1, 1), date(2024, 1, 5))
        mock_analyzer.analyze_task_data.assert_called_once_with(
            workbook_data.tasks, workbook_data.daily_tasks,
            workbook_data.communication_tasks, workbook_data.all_items)
        # ファイル名には実際にデータがあった最初と最後の日付を使う
        mock_writer.save_results.assert_called_once_with(
            ('results',), 'test_template.xlsx', 'test_output', date(2024, 1, 2), date(2024, 1, 4))

    def test_run_analysis_reloads_config(self, mocks):
        """起動後に設定ファイルを編集しても、分析時点の内容が使われること"""
        mock_load_config, mock_reader, _, _ = mocks
        mock_load_config.return_value = {'PATHS': {**PATHS, 'input_file_path': 'old.xlsx'}}
        analyzer = TaskAnalyzer()
        mock_load_config.return_value = {'PATHS': {**PATHS, 'input_file_path': 'new.xlsx'}}

        analyzer.run_analysis(date(2024, 1, 1), date(2024, 1, 5))

        assert mock_reader.read_workbook.call_args.args[0] == 'new.xlsx'

    def test_run_analysis_raises_error(self, mocks):
        """エラーは呼び出し側(GUI)で表示するため、そのまま送出すること"""
        _, mock_reader, _, _ = mocks
        mock_reader.read_workbook.side_effect = Exception('テストエラー')

        with pytest.raises(Exception, match='テストエラー'):
            TaskAnalyzer().run_analysis(date(2024, 1, 1), date(2024, 1, 5))
