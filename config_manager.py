import configparser
import os
import sys
from pathlib import Path
from typing import Final, List

def get_config_path() -> Path:
    # 実行ファイルのディレクトリを取得
    if getattr(sys, 'frozen', False):
        base_path = Path(sys._MEIPASS)
    else:
        base_path = Path(os.path.dirname(os.path.abspath(__file__)))
    return base_path / 'config.ini'

CONFIG_PATH = get_config_path()

class ConfigManager:
    def __init__(self, config_file: Path | str = CONFIG_PATH) -> None:
        self.config_file: Path = Path(config_file)
        self.config: configparser.ConfigParser = configparser.ConfigParser()
        self.load_config()

    def load_config(self) -> None:
        if not self.config_file.exists():
            raise FileNotFoundError(f"Config file not found: {self.config_file}")
        try:
            self.config.read(self.config_file, encoding='utf-8')
        except UnicodeDecodeError as e:
            raise OSError(f"Failed to load config: {e}") from e
   
    def get_processed_path(self) -> str:
        if 'Paths' not in self.config:
            return r"C:\Shinseikai\CSV2XL\processed"
        return self.config.get('Paths', 'processed_path', fallback=r"C:\Shinseikai\CSV2XL\processed")

    def set_processed_path(self, path: str) -> None:
        if 'Paths' not in self.config:
            self.config['Paths'] = {}
        self.config['Paths']['processed_path'] = path
        self.save_config()

    def get_font_size(self) -> int:
        if 'Appearance' not in self.config:
            return 9  # デフォルトのフォントサイズ
        return self.config.getint('Appearance', 'font_size', fallback=9)

    def set_font_size(self, size: int) -> None:
        if 'Appearance' not in self.config:
            self.config['Appearance'] = {}
        self.config['Appearance']['font_size'] = str(size)

    def get_window_size(self) -> tuple[int, int]:
        if 'Appearance' not in self.config:
            return 300, 300
        width = self.config.getint('Appearance', 'window_width', fallback=300)
        height = self.config.getint('Appearance', 'window_height', fallback=300)
        return width, height

    def set_window_size(self, width: int, height: int) -> None:
        if 'Appearance' not in self.config:
            self.config['Appearance'] = {}
        self.config['Appearance']['window_width'] = str(width)
        self.config['Appearance']['window_height'] = str(height)
        self.save_config()

    def save_config(self) -> None:
        try:
            with open(self.config_file, 'w', encoding='utf-8') as configfile:
                self.config.write(configfile)
        except (IOError, OSError) as e:
            raise OSError(f"Failed to load config: {e}") from e

    def _ensure_section(self, section: str) -> None:
        if section not in self.config:
            self.config[section] = {}
