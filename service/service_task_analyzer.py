from datetime import datetime
from utils.config_manager import load_config
from service.service_excel_reader import ExcelTaskReader
from service.service_data_analyzer import TaskDataAnalyzer
from service.service_excel_writer import ExcelResultWriter


class TaskAnalyzer:
    def run_analysis(self, start_date_str, end_date_str):
        try:
            # 起動後に設定ファイルが編集されていても反映されるよう、実行のたびに読み込む
            config = load_config()
            paths_config = config['PATHS']

            start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
            end_date = datetime.strptime(end_date_str, '%Y-%m-%d')

            tasks, daily_tasks, comm_tasks, all_items, actual_start_date_str, actual_end_date_str = ExcelTaskReader(config).read_workbook(
                paths_config['input_file_path'],
                start_date,
                end_date
            )

            actual_start_date = datetime.strptime(actual_start_date_str, '%Y%m%d')
            actual_end_date = datetime.strptime(actual_end_date_str, '%Y%m%d')

            analysis_results = TaskDataAnalyzer().analyze_task_data(
                tasks, daily_tasks, comm_tasks, all_items
            )

            output_file = ExcelResultWriter().save_results(
                analysis_results,
                paths_config['template_path'],
                paths_config['output_dir'],
                actual_start_date,
                actual_end_date
            )

            return True, f"分析が完了しました。結果は {output_file} に保存されました。"

        except Exception as e:
            return False, f"分析中にエラーが発生しました: {str(e)}"
