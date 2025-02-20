import tkinter as tk
from tkinter import ttk
from datetime import datetime
from tkcalendar import DateEntry
from config_manager import load_config, save_config
import configparser
import version

class TaskAnalyzerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f'業務分析アプリ v{version}')
        
        # 設定の読み込み
        self.config = load_config()
        
        # ウィンドウサイズの設定
        window_width = self.config.getint('Appearance', 'window_width')
        window_height = self.config.getint('Appearance', 'window_height')
        self.root.geometry(f"{window_width}x{window_height}")
        
        # フォントサイズの設定
        font_size = self.config.getint('Appearance', 'font_size')
        
        # メインフレーム
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 日付選択部分
        date_frame = ttk.LabelFrame(main_frame, text="分析期間", padding="5")
        date_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # 開始日
        ttk.Label(date_frame, text="開始日:").grid(row=0, column=0, padx=5, pady=5)
        self.start_date = DateEntry(date_frame, width=12, background='darkblue',
                                  foreground='white', borderwidth=2, year=2025,
                                  locale='ja_JP', date_pattern='yyyy/mm/dd')
        self.start_date.grid(row=0, column=1, padx=5, pady=5)
        
        # 終了日
        ttk.Label(date_frame, text="終了日:").grid(row=1, column=0, padx=5, pady=5)
        self.end_date = DateEntry(date_frame, width=12, background='darkblue',
                                foreground='white', borderwidth=2, year=2025,
                                locale='ja_JP', date_pattern='yyyy/mm/dd')
        self.end_date.grid(row=1, column=1, padx=5, pady=5)

        ttk.Button(main_frame, text="分析開始", command=self.start_analysis).grid(row=3, column=0, columnspan=2, pady=10)
        ttk.Button(main_frame, text="設定ファイル", command=self.open_config).grid(row=4, column=0, columnspan=2, pady=5)
        ttk.Button(main_frame, text="閉じる", command=self.root.quit).grid(row=5, column=0, columnspan=2, pady=5)
    
    def start_analysis(self):
        start_row = int(self.start_row.get())
        end_row = int(self.end_row.get())
        start_date = self.start_date.get_date()
        end_date = self.end_date.get_date()

        if 'Analysis' not in self.config:
            self.config.add_section('Analysis')
        
        self.config['Analysis'].update({
            'start_row': str(start_row),
            'end_row': str(end_row),
            'start_date': start_date.strftime('%Y-%m-%d'),
            'end_date': end_date.strftime('%Y-%m-%d')
        })
        
        save_config(self.config)
    
    def open_config(self):
        # 設定ファイルを開く処理
        pass

def main():
    root = tk.Tk()
    app = TaskAnalyzerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
