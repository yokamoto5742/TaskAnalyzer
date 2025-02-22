import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkcalendar import DateEntry
from datetime import datetime
import subprocess

from config_manager import load_config, save_config
from version import VERSION
from service_analysis import TaskAnalyzer


class TaskAnalyzerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f'業務分析アプリ v{VERSION}')

        # 設定の読み込み
        self.config = load_config()

        # ウィンドウサイズの設定
        window_width = self.config.getint('Appearance', 'window_width')
        window_height = self.config.getint('Appearance', 'window_height')
        self.root.geometry(f"{window_width}x{window_height}")

        # フォントサイズの設定
        font_size = self.config.getint('Appearance', 'font_size')

        self._setup_gui()

    def _setup_gui(self):
        """GUIコンポーネントの初期化とレイアウト設定"""
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 日付選択部分
        self._setup_date_frame(main_frame)

        # ボタン類
        self._setup_buttons(main_frame)

    def _setup_date_frame(self, parent):
        """日付選択部分のセットアップ"""
        date_frame = ttk.LabelFrame(parent, text="分析期間", padding="5")
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

    def _setup_buttons(self, parent):
        """ボタン類のセットアップ"""
        ttk.Button(parent, text="分析開始", command=self.start_analysis).grid(
            row=3, column=0, columnspan=2, pady=10)
        ttk.Button(parent, text="設定ファイル", command=self.open_config).grid(
            row=4, column=0, columnspan=2, pady=5)
        ttk.Button(parent, text="閉じる", command=self.root.quit).grid(
            row=5, column=0, columnspan=2, pady=5)

    def start_analysis(self):
        """分析開始ボタンのコールバック"""
        try:
            # 日付の取得と妥当性チェック
            start_date = self.start_date.get_date()
            end_date = self.end_date.get_date()

            if start_date > end_date:
                messagebox.showerror("エラー", "開始日が終了日より後の日付になっています。")
                return

            # 設定の更新
            if 'Analysis' not in self.config:
                self.config.add_section('Analysis')

            self.config['Analysis'].update({
                'start_date': start_date.strftime('%Y-%m-%d'),
                'end_date': end_date.strftime('%Y-%m-%d')
            })
            save_config(self.config)

            # 分析処理の実行
            analyzer = TaskAnalyzer()
            self.root.config(cursor="wait")  # カーソルを待機中に変更
            self.root.update()

            try:
                start_date_str = start_date.strftime('%Y-%m-%d')
                end_date_str = end_date.strftime('%Y-%m-%d')
                success = analyzer.run_analysis(start_date_str, end_date_str)
                if success:
                    messagebox.showinfo("完了", "分析が完了しました。")
                else:
                    messagebox.showerror("エラー", "分析中にエラーが発生しました。")
            finally:
                self.root.config(cursor="")  # カーソルを元に戻す

        except Exception as e:
            messagebox.showerror("エラー", f"予期せぬエラーが発生しました：\n{str(e)}")
            print(f"エラーの詳細: {e}")

    def open_config(self):
        """設定ファイルを開くボタンのコールバック"""
        # 設定ファイルのパスを取得
        config_path = self.config.get('PATHS', 'config_path')

        try:
            subprocess.Popen(['notepad.exe', config_path])
        except Exception as e:
            messagebox.showerror("エラー", f"設定ファイルを開けませんでした：\n{str(e)}")
            print(f"エラーの詳細: {e}")
