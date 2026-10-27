import pytest
import polars as pl
from service.service_data_analyzer import TaskDataAnalyzer, summarize


@pytest.fixture
def sample_tasks():
    return [
        {'date': '2024-01-01', 'content': 'クラーク業務A', 'minutes': 30},
        {'date': '2024-01-01', 'content': 'クラーク業務B', 'minutes': 45},
        {'date': '2024-01-02', 'content': 'クラーク業務A', 'minutes': 25},
        {'date': '2024-01-02', 'content': '会議', 'minutes': 60},
        {'date': '2024-01-03', 'content': '資料作成', 'minutes': 90},
    ]


@pytest.fixture
def sample_daily_tasks():
    return [
        {'date': '2024-01-01', 'content': '毎日タスクA', 'minutes': 10},
        {'date': '2024-01-01', 'content': '毎日タスクB', 'minutes': 15},
        {'date': '2024-01-02', 'content': '毎日タスクA', 'minutes': 10},
        {'date': '2024-01-03', 'content': '毎日タスクB', 'minutes': 15},
    ]


@pytest.fixture
def sample_communication_tasks():
    return [
        {'date': '2024-01-01', 'name': '田中', 'content': '打合せ', 'minutes': 30},
        {'date': '2024-01-01', 'name': '佐藤', 'content': 'レビュー', 'minutes': 45},
        {'date': '2024-01-02', 'name': '田中', 'content': '打合せ', 'minutes': 25},
        {'date': '2024-01-02', 'name': '鈴木', 'content': '相談', 'minutes': 15},
        {'date': '2024-01-03', 'name': '佐藤', 'content': 'レビュー', 'minutes': 30},
    ]


@pytest.fixture
def sample_all_items():
    return [
        {'date': '2024-01-01', 'content': 'クラーク業務A', 'minutes': 30},
        {'date': '2024-01-01', 'content': '毎日タスクA', 'minutes': 10},
        {'date': '2024-01-02', 'content': '会議', 'minutes': 60},
        {'date': '2024-01-03', 'content': '資料作成', 'minutes': 90},
    ]


def test_summarize(sample_tasks):
    result = summarize(pl.DataFrame(sample_tasks), ['content'])

    assert result.columns == ['content', 'total_minutes', 'total_hours', 'frequency']
    assert result['content'].to_list() == ['資料作成', '会議', 'クラーク業務A', 'クラーク業務B']
    clerk_a = result.filter(pl.col('content') == 'クラーク業務A')
    assert clerk_a['total_minutes'][0] == 55
    assert clerk_a['frequency'][0] == 2


def test_summarize_by_two_keys(sample_communication_tasks):
    result = summarize(pl.DataFrame(sample_communication_tasks), ['content', 'name'])

    assert result.columns[:2] == ['content', 'name']
    tanaka_meeting = result.filter((pl.col('name') == '田中') & (pl.col('content') == '打合せ'))
    assert tanaka_meeting['total_minutes'][0] == 55


class TestTaskDataAnalyzer:
    def test_analyze_task_data_with_empty_lists(self, sample_tasks):
        """コミュニケーション・デイリータスクが 0 件の期間でも集計できること"""
        results = TaskDataAnalyzer().analyze_task_data(sample_tasks, [], [], sample_tasks)

        _, _, daily_tasks_agg, communication_by_name, communication_by_content, _ = results
        assert len(daily_tasks_agg) == 0
        assert len(communication_by_name) == 0
        assert communication_by_content.columns == [
            'content', 'name', 'total_minutes', 'total_hours', 'frequency']

    def test_total_hours_is_not_truncated(self):
        """90 分は 1 時間に切り捨てず 1.5 時間になること"""
        tasks = [{'date': '2024-01-01', 'content': '会議', 'minutes': 90}]

        result = TaskDataAnalyzer().analyze_task_data(tasks, tasks, [], tasks)[1]

        assert result['total_hours'][0] == 1.5

    def test_analyze_task_data(self, sample_tasks, sample_daily_tasks, sample_communication_tasks, sample_all_items):
        # テスト準備
        analyzer = TaskDataAnalyzer()
        
        # テスト実行
        results = analyzer.analyze_task_data(
            sample_tasks, sample_daily_tasks, sample_communication_tasks, sample_all_items
        )
        
        # 検証
        clerk_tasks, non_clerk_tasks, daily_tasks_agg, communication_by_name, communication_by_content, all_items_summary = results
        
        # クラーク業務の確認
        assert len(clerk_tasks) == 2
        
        # クラーク業務以外の確認
        assert len(non_clerk_tasks) == 2
        
        # デイリータスクの確認
        assert len(daily_tasks_agg) == 2
        
        # コミュニケーション（名前別）の確認
        assert len(communication_by_name) == 3
        
        # コミュニケーション（内容別）の確認
        assert len(communication_by_content) == 3
        assert communication_by_content.columns == [
            'content', 'name', 'total_minutes', 'total_hours', 'frequency']
        assert communication_by_content['name'].to_list() == sorted(communication_by_content['name'].to_list())
        
        # 全項目の確認
        assert len(all_items_summary) == 4
