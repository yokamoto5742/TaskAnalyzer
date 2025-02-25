from datetime import datetime
from config_manager import load_config
from service_excel_reader import ExcelTaskReader
from service_data_analyzer import TaskDataAnalyzer
from service_excel_writer import ExcelResultWriter


class TaskAnalyzer:
    """
    業務分析を行うメインクラス
    元のservice_analysis.pyの代わりとなるクラス
    """
    def __init__(self):
        self.config = load_config()
        self.paths_config = self.config['PATHS']
        self.reader = ExcelTaskReader(self.config)
        self.analyzer = TaskDataAnalyzer()
        self.writer = ExcelResultWriter()

    def run_analysis(self, start_date_str, end_date_str):
        """
        指定された期間の業務分析を実行する
        
        Parameters:
        -----------
        start_date_str : str
            分析開始日 (YYYY-MM-DD形式)
        end_date_str : str
            分析終了日 (YYYY-MM-DD形式)
            
        Returns:
        --------
        tuple
            (success, error_message)
        """
        try:
            # 文字列をdatetimeオブジェクトに変換
            start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
            end_date = datetime.strptime(end_date_str, '%Y-%m-%d')

            # データの読み込み
            tasks, daily_tasks, comm_tasks, all_items, _, _ = self.reader.read_workbook(
                self.paths_config['input_file_path'],
                start_date,
                end_date
            )

            # データの分析
            analysis_results = self.analyzer.analyze_task_data(
                tasks, daily_tasks, comm_tasks, all_items
            )

            # 結果の保存
            self.writer.save_results(
                analysis_results,
                self.paths_config['template_path'],
                self.paths_config['output_dir'],
                start_date,
                end_date
            )

            return True, None

        except ValueError as ve:
            return False, f"日付の形式が正しくありません: {str(ve)}"
        except Exception as e:
            return False, f"分析中にエラーが発生しました: {str(e)}"
