import configparser
from unittest.mock import Mock, patch

import pytest

from app.path_settings_dialog import PathSettingsDialog


@pytest.fixture
def dialog():
    config = configparser.ConfigParser()
    config['PATHS'] = {'input_file_path': 'a.xlsx', 'template_path': 't.xlsx', 'output_dir': 'out'}
    with patch('app.path_settings_dialog.tk.Toplevel'), \
            patch('app.path_settings_dialog.tk.StringVar',
                  side_effect=lambda value: Mock(get=Mock(return_value=value), set=Mock())), \
            patch('app.path_settings_dialog.ttk'), \
            patch('app.path_settings_dialog.load_config', return_value=config):
        yield PathSettingsDialog(Mock())


def test_save_updates_paths_only(dialog):
    """保存すると PATHS だけが更新され、他のセクションは残ること"""
    dialog.path_vars['input_file_path'].get.return_value = ' new.xlsx '
    current = configparser.ConfigParser()
    current['PATHS'] = {'input_file_path': 'a.xlsx'}
    current['Analysis'] = {'start_row': '4'}

    with patch('app.path_settings_dialog.load_config', return_value=current), \
            patch('app.path_settings_dialog.save_config') as mock_save:
        dialog.save()

    saved = mock_save.call_args.args[0]
    assert saved['PATHS']['input_file_path'] == 'new.xlsx'
    assert saved['PATHS']['output_dir'] == 'out'
    assert saved['Analysis']['start_row'] == '4'
    dialog.dialog.destroy.assert_called_once()


def test_choose_file_sets_path(dialog):
    """選んだファイルのパスが Windows 形式で入ること"""
    with patch('app.path_settings_dialog.filedialog.askopenfilename', return_value='C:/x/a.xlsx'):
        dialog._choose_file('input_file_path')

    dialog.path_vars['input_file_path'].set.assert_called_once_with('C:\\x\\a.xlsx')


def test_choose_directory_cancel_keeps_path(dialog):
    """キャンセルしたら入力欄を変えないこと"""
    with patch('app.path_settings_dialog.filedialog.askdirectory', return_value=''):
        dialog._choose_directory('output_dir')

    dialog.path_vars['output_dir'].set.assert_not_called()
