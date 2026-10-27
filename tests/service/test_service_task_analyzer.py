from unittest.mock import MagicMock, patch
from datetime import datetime
from service.service_task_analyzer import TaskAnalyzer


class TestTaskAnalyzer:
    @patch('service.service_task_analyzer.load_config')
    @patch('service.service_task_analyzer.ExcelTaskReader')
    @patch('service.service_task_analyzer.TaskDataAnalyzer')
    @patch('service.service_task_analyzer.ExcelResultWriter')
    def test_run_analysis_success(self, mock_writer_class, mock_analyzer_class, mock_reader_class, mock_load_config):
        # モックの設定
        mock_config = {
            'PATHS': {
                'input_file_path': 'test_input.xlsx',
                'template_path': 'test_template.xlsx',
                'output_dir': 'test_output'
            }
        }
        mock_load_config.return_value = mock_config
        
        # リーダーのモック
        mock_reader = MagicMock()
        mock_reader_class.return_value = mock_reader
        mock_reader.read_workbook.return_value = (
            [{'content': 'タスク1', 'minutes': 30}],  # tasks
            [{'content': '日次タスク1', 'minutes': 10}],  # daily_tasks
            [{'name': '田中', 'content': '打合せ', 'minutes': 30}],  # comm_tasks
            [{'content': 'すべて1', 'minutes': 30}],  # all_items
            '20240101',  # actual_start_date_str
            '20240105'   # actual_end_date_str
        )
        
        # アナライザーのモック
        mock_analyzer = MagicMock()
        mock_analyzer_class.return_value = mock_analyzer
        mock_analyzer.analyze_task_data.return_value = (
            'clerk_tasks',
            'non_clerk_tasks',
            'daily_tasks_agg',
            'communication_by_name',
            'communication_by_content',
            'all_items_summary'
        )
        
        # ライターのモック
        mock_writer = MagicMock()
        mock_writer_class.return_value = mock_writer
        mock_writer.save_results.return_value = 'output_file_path.xlsx'
        
        # テスト実行
        analyzer = TaskAnalyzer()
        success, message = analyzer.run_analysis('2024-01-01', '2024-01-05')
        
        # 検証
        assert success is True
        assert '分析が完了しました' in message
        assert 'output_file_path.xlsx' in message
        
        # 各メソッドが正しく呼び出されたことを検証
        mock_reader.read_workbook.assert_called_once_with(
            'test_input.xlsx',
            datetime(2024, 1, 1),
            datetime(2024, 1, 5)
        )
        
        mock_analyzer.analyze_task_data.assert_called_once_with(
            [{'content': 'タスク1', 'minutes': 30}],
            [{'content': '日次タスク1', 'minutes': 10}],
            [{'name': '田中', 'content': '打合せ', 'minutes': 30}],
            [{'content': 'すべて1', 'minutes': 30}]
        )
        
        mock_writer.save_results.assert_called_once_with(
            (
                'clerk_tasks',
                'non_clerk_tasks',
                'daily_tasks_agg',
                'communication_by_name',
                'communication_by_content',
                'all_items_summary'
            ),
            'test_template.xlsx',
            'test_output',
            datetime(2024, 1, 1),
            datetime(2024, 1, 5)
        )

    @patch('service.service_task_analyzer.load_config')
    @patch('service.service_task_analyzer.ExcelTaskReader')
    def test_run_analysis_no_data_message(self, mock_reader_class, mock_load_config):
        """データがないときに「日付の形式」のエラーと表示されないこと"""
        mock_load_config.return_value = {'PATHS': {'input_file_path': 'test_input.xlsx'}}
        mock_reader_class.return_value.read_workbook.side_effect = ValueError('指定された期間内のデータがありません')

        success, message = TaskAnalyzer().run_analysis('2024-01-01', '2024-01-05')

        assert success is False
        assert '指定された期間内のデータがありません' in message
        assert '日付の形式' not in message

    @patch('service.service_task_analyzer.load_config')
    @patch('service.service_task_analyzer.ExcelTaskReader')
    @patch('service.service_task_analyzer.TaskDataAnalyzer')
    @patch('service.service_task_analyzer.ExcelResultWriter')
    def test_run_analysis_reloads_config(self, _mock_writer, _mock_analyzer, mock_reader_class, mock_load_config):
        """起動後に設定ファイルを編集しても、分析時点の内容が使われること"""
        mock_load_config.return_value = {'PATHS': {'input_file_path': 'old.xlsx'}}
        analyzer = TaskAnalyzer()
        mock_load_config.return_value = {
            'PATHS': {'input_file_path': 'new.xlsx', 'template_path': 't.xlsx', 'output_dir': 'out'}}
        mock_reader_class.return_value.read_workbook.return_value = ([], [], [], [], '20240101', '20240105')

        analyzer.run_analysis('2024-01-01', '2024-01-05')

        assert mock_reader_class.return_value.read_workbook.call_args.args[0] == 'new.xlsx'

    @patch('service.service_task_analyzer.load_config')
    @patch('service.service_task_analyzer.ExcelTaskReader')
    def test_run_analysis_general_exception(self, mock_reader_class, mock_load_config):
        # モックの設定
        mock_config = {
            'PATHS': {
                'input_file_path': 'test_input.xlsx',
                'template_path': 'test_template.xlsx',
                'output_dir': 'test_output'
            }
        }
        mock_load_config.return_value = mock_config
        
        # リーダーが例外を発生させるようにモックする
        mock_reader = MagicMock()
        mock_reader_class.return_value = mock_reader
        mock_reader.read_workbook.side_effect = Exception('テストエラー')
        
        # テスト実行
        analyzer = TaskAnalyzer()
        success, message = analyzer.run_analysis('2024-01-01', '2024-01-05')
        
        # 検証
        assert success is False
        assert '分析中にエラーが発生しました' in message
        assert 'テストエラー' in message
