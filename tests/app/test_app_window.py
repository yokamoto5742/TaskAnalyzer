import pytest
from unittest.mock import Mock, patch
from datetime import date
import configparser
from app.app_window import CONFIG_PATH, TaskAnalyzerGUI
from app import __version__ as VERSION
from app.constants import INVALID_DATE_MESSAGE


@pytest.fixture
def mock_tk():
    with patch('app.app_window.tk.Tk') as mock_tk_class:
        root = Mock()
        root.title = Mock()
        root.geometry = Mock()
        root.quit = Mock()
        mock_tk_class.return_value = root
        yield root


@pytest.fixture
def mock_date_entry():
    with patch('app.app_window.DateEntry') as mock_date_entry_class:
        date_entry = Mock()
        date_entry.get_date = Mock()
        mock_date_entry_class.return_value = date_entry
        yield date_entry


@pytest.fixture
def mock_config():
    config = configparser.ConfigParser()
    original_config = configparser.ConfigParser()

    config['Appearance'] = {
        'window_width': '800',
        'window_height': '600'
    }
    config['PATHS'] = {}
    config['Analysis'] = {}

    for section in config.sections():
        if not original_config.has_section(section):
            original_config.add_section(section)
        for key, value in config[section].items():
            original_config[section][key] = value

    with patch('app.app_window.load_config', return_value=config), \
            patch('app.app_window.save_config'):
        yield config
        restore_config(config, original_config)


def restore_config(config, original_config):
    """configを元の状態に復元するヘルパーメソッド"""
    for section in config.sections():
        config.remove_section(section)
    for section in original_config.sections():
        if not config.has_section(section):
            config.add_section(section)
        for key, value in original_config[section].items():
            config[section][key] = value



@pytest.fixture
def mock_analyzer():
    with patch('app.app_window.TaskAnalyzer') as mock_task_analyzer:
        analyzer = Mock()
        mock_task_analyzer.return_value = analyzer
        yield analyzer


@pytest.fixture
def mock_messagebox():
    with patch('app.app_window.messagebox') as mock_msg:
        # showerrorのデフォルト実装を設定
        mock_msg.showerror = Mock()
        yield mock_msg


@pytest.fixture
def gui(mock_tk, mock_config, mock_analyzer, mock_date_entry):
    with patch('app.app_window.ttk.Frame'), \
            patch('app.app_window.ttk.LabelFrame'), \
            patch('app.app_window.ttk.Label'), \
            patch('app.app_window.ttk.Button'), \
            patch('app.app_window.ttk.Spinbox'), \
            patch('app.app_window.ttk.Radiobutton'), \
            patch('app.app_window.ttk.Combobox'), \
            patch('app.app_window.tk.StringVar'), \
            patch('app.app_window.tk.IntVar'):
        gui = TaskAnalyzerGUI(mock_tk)
        gui.year_start_month_var = Mock(get=Mock(return_value=1))
        yield gui


def test_init(gui, mock_tk):
    """初期化のテスト"""
    mock_tk.title.assert_called_with(f'業務分析 v{VERSION}')


def test_init_restores_saved_dates(mock_tk, mock_config, mock_analyzer, mock_date_entry):
    """保存済みの期間が起動時に入ること"""
    mock_config['Analysis'] = {'start_date': '2025-04-01', 'end_date': '2025-06-30'}

    with patch('app.app_window.ttk.Frame'), \
            patch('app.app_window.ttk.LabelFrame'), \
            patch('app.app_window.ttk.Label'), \
            patch('app.app_window.ttk.Button'), \
            patch('app.app_window.ttk.Spinbox'), \
            patch('app.app_window.ttk.Radiobutton'), \
            patch('app.app_window.ttk.Combobox'), \
            patch('app.app_window.tk.StringVar') as mock_string_var, \
            patch('app.app_window.tk.IntVar'):
        TaskAnalyzerGUI(mock_tk)

    assert [c.args[0] for c in mock_date_entry.set_date.call_args_list] == [
        date(2025, 4, 1), date(2025, 6, 30)]
    mock_string_var.assert_any_call(value='2025')


def test_apply_period(gui, mock_date_entry):
    """期間を選ぶと開始日と終了日が入ること"""
    gui.period_var = Mock(get=Mock(return_value='上半期'))
    gui.year_var = Mock(get=Mock(return_value='2025'))
    gui.year_start_month_var = Mock(get=Mock(return_value=4))

    gui._apply_period()

    assert [c.args[0] for c in mock_date_entry.set_date.call_args_list[-2:]] == [
        date(2025, 4, 1), date(2025, 9, 30)]


def test_analysis_error(gui, mock_config, mock_analyzer, mock_date_entry, mock_messagebox):
    """分析エラー時のテスト"""
    error_message = "テストエラー"
    mock_analyzer.run_analysis.return_value = (False, error_message)

    mock_date_entry.get.side_effect = ['2025/02/01', '2025/02/28']

    with patch('app.app_window.save_config'):
        gui.start_analysis()

    mock_analyzer.run_analysis.assert_called_once_with('2025-02-01', '2025-02-28')
    mock_messagebox.showerror.assert_called_with("エラー", error_message)
    assert mock_config['Analysis']['year_start_month'] == '1'


def test_analysis_invalid_date(gui, mock_analyzer, mock_date_entry, mock_messagebox):
    """yyyy/mm/dd 以外の入力はエラーになること"""
    mock_date_entry.get.side_effect = ['2025-02-01', '2025/02/28']

    gui.start_analysis()

    mock_analyzer.run_analysis.assert_not_called()
    mock_messagebox.showerror.assert_called_with("エラー", INVALID_DATE_MESSAGE)


def test_open_config(gui, mock_config):
    """設定ファイルを開く機能のテスト"""
    with patch('app.app_window.subprocess.Popen') as mock_popen:
        gui.open_config()
        mock_popen.assert_called_once_with(['notepad.exe', CONFIG_PATH])


def test_open_config_error(gui, mock_config):
    """設定ファイルを開く際のエラーテスト"""
    with patch('app.app_window.subprocess.Popen', side_effect=Exception("テストエラー")):
        with patch('app.app_window.messagebox.showerror') as mock_error:
            gui.open_config()
            mock_error.assert_called_with("エラー", "設定ファイルを開けませんでした：\nテストエラー")
