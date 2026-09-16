import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.ticker import AutoMinorLocator
import pandas as pd
import numpy as np
import os
import csv
import re
import itertools
import io
import json

# フォントが見つからない場合のエラーを回避する安全な設定
try:
    import matplotlib.font_manager as fm
    fonts = [f.name for f in fm.fontManager.ttflist]
    if 'MS Gothic' in fonts:
        plt.rcParams['font.family'] = 'MS Gothic'
    elif 'Meiryo' in fonts:
        plt.rcParams['font.family'] = 'Meiryo'
    elif 'Yu Gothic' in fonts:
        plt.rcParams['font.family'] = 'Yu Gothic'
    else:
        plt.rcParams['font.family'] = 'sans-serif' # 見つからない場合はデフォルト
except Exception:
    pass

class GraphApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV/Excel/LTspice Graph Generator (Live Preview)")
        self.root.geometry("1400x900")
        
        # テーマ設定
        style = ttk.Style()
        if 'clam' in style.theme_names():
            style.theme_use('clam')
            
        # 内部データの保持
        self.file_paths = []
        self.raw_dfs = []
        self.loaded_file_labels = []
        self.file_settings = {} # { filepath: {'x_prefix': tk.StringVar, 'y_prefix': tk.StringVar} }
        self.clipboard_source_id = "__clipboard__"
        self.source_default_styles = {}
        self.series_keys = []
        self.series_color_palette = ["自動", "blue", "orange", "green", "red", "purple", "brown", "pink", "gray", "olive", "cyan", "black"]
        self.marker_options = ["o", "s", "^", "v", "D", "x", "+", "*", ".", "なし"]
        self.line_style_options = ["実線", "破線", "点線", "一点鎖線", "なし"]
        
        self._draw_timer = None # デバウンス用タイマー
        
        # --- UI用変数 ---
        self.load_mode = tk.StringVar(value="replace")
        self.header_row = tk.StringVar(value="")
        self.x_col_index = tk.IntVar(value=0)
        
        self.x_min = tk.StringVar(); self.x_max = tk.StringVar()
        self.y_min = tk.StringVar(); self.y_max = tk.StringVar()
        self.x_label = tk.StringVar(value="x軸"); self.y_label = tk.StringVar(value="y軸")
        
        self.show_y_minmax = tk.BooleanVar(value=False)
        self.show_annotations = tk.BooleanVar(value=True)
        self.x_log_scale = tk.BooleanVar(value=False)
        
        # --- デザイン・サイズ設定用変数 ---
        # 全体的に文字を見やすくするため、デフォルトを大きめに設定
        self.base_font_size = tk.IntVar(value=14)
        self.line_width = tk.DoubleVar(value=2.0)
        self.marker_size = tk.IntVar(value=40)
        
        self.status_message = tk.StringVar(value="準備完了: ファイルを選択してください")
        self.preview_source_var = tk.StringVar(value="")
        
        # 吹き出し・曲線設定 (読み込んだ系列数に合わせて動的に増やす)
        self.show_series_vars = []
        self.col_name_vars = []
        self.anno_text_vars = []
        self.anno_style_vars = []
        self.anno_mode_vars = []
        self.anno_status_vars = []
        self.anno_click_positions = []
        self.series_color_vars = []
        self.marker_style_vars = []
        self.line_style_vars = []
        
        self.prefixes = ["", "p", "n", "μ", "m", "k", "M", "G"]
        
        self.create_widgets()
        self.bind_events()

    def create_widgets(self):
        # 画面全体を左右に分割するメインフレーム
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- 左パネル (コントロール群) ---
        # widthを固定して、UIが潰れて見えなくなるのを防ぐ
        left_container = ttk.Frame(main_frame, width=480)
        left_container.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)
        left_container.pack_propagate(False) # 中身に合わせて縮まないようにする
        
        self.control_canvas = tk.Canvas(left_container, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(left_container, orient=tk.VERTICAL, command=self.control_canvas.yview)
        self.control_frame = ttk.Frame(self.control_canvas, padding=10)
        
        self.control_canvas.configure(yscrollcommand=self.scrollbar.set)
        self.control_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.canvas_window = self.control_canvas.create_window((0, 0), window=self.control_frame, anchor="nw")
        
        # スクロール領域の自動調整
        self.control_frame.bind("<Configure>", lambda e: self.control_canvas.configure(scrollregion=self.control_canvas.bbox("all")))
        self.control_canvas.bind("<Configure>", lambda e: self.control_canvas.itemconfig(self.canvas_window, width=e.width))

        # 1. ファイル選択
        file_frame = ttk.LabelFrame(self.control_frame, text="1. データ選択 (CSV / txt / xlsx / xls / クリップボード)", padding=10)
        file_frame.pack(fill=tk.X, pady=5)
        
        button_frame = ttk.Frame(file_frame)
        button_frame.pack(fill=tk.X, pady=5)
        ttk.Button(button_frame, text="ファイルを選択", command=self.select_file).pack(side=tk.LEFT)
        ttk.Button(button_frame, text="クリップボードから読み込み", command=self.load_clipboard_data).pack(side=tk.LEFT, padx=(8, 0))
        
        load_mode_frame = ttk.Frame(file_frame)
        load_mode_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(load_mode_frame, text="読み込み方法:").pack(side=tk.LEFT)
        ttk.Radiobutton(load_mode_frame, text="新しいグラフに置き換え", variable=self.load_mode, value="replace").pack(side=tk.LEFT, padx=(8, 0))
        ttk.Radiobutton(load_mode_frame, text="現在のグラフに追加", variable=self.load_mode, value="append").pack(side=tk.LEFT, padx=(8, 0))
        
        # 動的にファイル設定を生成するフレーム
        self.file_settings_container = ttk.Frame(file_frame)
        self.file_settings_container.pack(fill=tk.X, pady=5)
        
        csv_setting_frame = ttk.Frame(file_frame)
        csv_setting_frame.pack(fill=tk.X, pady=5)
        ttk.Label(csv_setting_frame, text="ヘッダー行番号 (空欄で自動):").pack(side=tk.LEFT)
        ttk.Entry(csv_setting_frame, textvariable=self.header_row, width=5).pack(side=tk.LEFT, padx=5)
        ttk.Label(csv_setting_frame, text="X軸列番号 (0始まり):").pack(side=tk.LEFT, padx=(10, 0))
        ttk.Entry(csv_setting_frame, textvariable=self.x_col_index, width=5).pack(side=tk.LEFT, padx=5)

        # 2. 軸設定
        axis_frame = ttk.LabelFrame(self.control_frame, text="2. 軸・表示設定", padding=10)
        axis_frame.pack(fill=tk.X, pady=5)

        ttk.Label(axis_frame, text="設定項目").grid(row=0, column=0, sticky="w")
        ttk.Label(axis_frame, text="X軸 (横)").grid(row=0, column=1)
        ttk.Label(axis_frame, text="Y軸 (縦)").grid(row=0, column=2)

        ttk.Label(axis_frame, text="最小値 (空で自動):").grid(row=1, column=0, sticky="e", pady=2)
        ttk.Entry(axis_frame, textvariable=self.x_min, width=10).grid(row=1, column=1, padx=5)
        ttk.Entry(axis_frame, textvariable=self.y_min, width=10).grid(row=1, column=2, padx=5)

        ttk.Label(axis_frame, text="最大値 (空で自動):").grid(row=2, column=0, sticky="e", pady=2)
        ttk.Entry(axis_frame, textvariable=self.x_max, width=10).grid(row=2, column=1, padx=5)
        ttk.Entry(axis_frame, textvariable=self.y_max, width=10).grid(row=2, column=2, padx=5)

        ttk.Label(axis_frame, text="軸ラベル:").grid(row=3, column=0, sticky="e", pady=2)
        ttk.Entry(axis_frame, textvariable=self.x_label, width=15).grid(row=3, column=1, padx=5)
        ttk.Entry(axis_frame, textvariable=self.y_label, width=15).grid(row=3, column=2, padx=5)

        ttk.Checkbutton(axis_frame, text="Y最小値・最大値を表示", variable=self.show_y_minmax).grid(row=4, column=0, columnspan=2, sticky="w", pady=(5,0))
        ttk.Checkbutton(axis_frame, text="X軸を対数(Log)スケール", variable=self.x_log_scale).grid(row=5, column=0, columnspan=2, sticky="w", pady=2)

        # 3. 吹き出し設定
        anno_frame = ttk.LabelFrame(self.control_frame, text="3. 吹き出し・曲線設定", padding=10)
        anno_frame.pack(fill=tk.X, pady=5)
        
        ttk.Checkbutton(anno_frame, text="曲線の吹き出しを表示", variable=self.show_annotations).pack(anchor="w", pady=(0, 6))
        self.anno_rows_container = ttk.Frame(anno_frame)
        self.anno_rows_container.pack(fill=tk.X)
        ttk.Label(self.anno_rows_container, text="データを読み込むと系列設定が表示されます", foreground="gray").pack(anchor="w")
        
        # 4. デザイン・サイズ設定
        design_frame = ttk.LabelFrame(self.control_frame, text="4. デザイン・サイズ設定", padding=10)
        design_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(design_frame, text="文字サイズ:").grid(row=0, column=0, sticky="e", pady=2)
        ttk.Spinbox(design_frame, textvariable=self.base_font_size, from_=8, to=36, width=5).grid(row=0, column=1, padx=5, sticky="w")
        
        ttk.Label(design_frame, text="線の太さ:").grid(row=0, column=2, sticky="e", pady=2)
        ttk.Spinbox(design_frame, textvariable=self.line_width, from_=0.5, to=10.0, increment=0.5, width=5).grid(row=0, column=3, padx=5, sticky="w")
        
        ttk.Label(design_frame, text="点の大きさ:").grid(row=0, column=4, sticky="e", pady=2)
        ttk.Spinbox(design_frame, textvariable=self.marker_size, from_=5, to=200, increment=5, width=5).grid(row=0, column=5, padx=5, sticky="w")

        # 5. 保存・出力
        output_frame = ttk.LabelFrame(self.control_frame, text="5. 保存・出力", padding=10)
        output_frame.pack(fill=tk.X, pady=5)
        ttk.Button(output_frame, text="画像/PDF/SVG出力", command=self.export_graph).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(output_frame, text="設定保存", command=self.save_settings).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(output_frame, text="設定読み込み", command=self.load_settings).pack(side=tk.LEFT)

        # 6. データプレビュー
        preview_frame = ttk.LabelFrame(self.control_frame, text="6. データプレビュー", padding=10)
        preview_frame.pack(fill=tk.X, pady=5)
        preview_select_frame = ttk.Frame(preview_frame)
        preview_select_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(preview_select_frame, text="表示データ:").pack(side=tk.LEFT)
        self.preview_source_combo = ttk.Combobox(preview_select_frame, textvariable=self.preview_source_var, width=28, state="readonly")
        self.preview_source_combo.pack(side=tk.LEFT, padx=5)
        ttk.Button(preview_select_frame, text="更新", command=self.update_data_preview).pack(side=tk.LEFT)
        self.preview_tree = ttk.Treeview(preview_frame, height=7, show="headings")
        self.preview_tree.pack(fill=tk.X)
        preview_scroll = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.preview_tree.xview)
        preview_scroll.pack(fill=tk.X)
        self.preview_tree.configure(xscrollcommand=preview_scroll.set)

        # ステータス表示領域
        status_label = ttk.Label(self.control_frame, textvariable=self.status_message, foreground="darkgreen", wraplength=420)
        status_label.pack(fill=tk.X, pady=10)

        # --- 右パネル (プレビュー) ---
        self.graph_frame = ttk.Frame(main_frame, padding=5)
        self.graph_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.fig, self.ax = plt.subplots(figsize=(8, 6), dpi=100)
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.graph_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.graph_frame)
        self.toolbar.update()
        
        # 初期状態の空グラフを描画
        self.ax.text(0.5, 0.5, "データを読み込むとここにグラフが表示されます", 
                     horizontalalignment='center', verticalalignment='center', 
                     transform=self.ax.transAxes, color='gray')
        self.canvas.draw()
        
        # クリックイベントのバインド
        self.fig.canvas.mpl_connect('button_press_event', self.on_canvas_click)

    def bind_events(self):
        """UIの変更を監視し、リアルタイム更新をスケジュールする"""
        vars_to_bind = [
            self.header_row, self.x_col_index, self.x_min, self.x_max, 
            self.y_min, self.y_max, self.x_label, self.y_label,
            self.show_y_minmax, self.show_annotations, self.x_log_scale,
            self.base_font_size, self.line_width, self.marker_size
        ]
        
        for var in vars_to_bind:
            var.trace_add("write", self.schedule_update)
        self.x_col_index.trace_add("write", lambda *args: self.update_column_info(reset_annotations=True))
        self.preview_source_var.trace_add("write", lambda *args: self.update_data_preview())
        self.root.bind_all("<Control-v>", self.handle_ctrl_v)

    def schedule_update(self, *args):
        """入力ラグを防ぐためのデバウンス処理（0.3秒待ってから描画）"""
        if self._draw_timer is not None:
            self.root.after_cancel(self._draw_timer)
        self._draw_timer = self.root.after(300, self.update_plot)

    def select_file(self):
        filenames = filedialog.askopenfilenames(
            filetypes=[("Data Files", "*.csv *.txt *.xlsx *.xls"), ("All Files", "*.*")]
        )
        if filenames:
            if self.load_mode.get() == "append":
                self.append_file_data(filenames)
            else:
                self.file_paths = list(filenames)
                self.source_default_styles = {}
                self.load_all_data()

    def load_clipboard_data(self):
        try:
            text = self.root.clipboard_get()
        except tk.TclError:
            message = "クリップボードにテキストデータがありません。Excelでセル範囲をコピーしてから実行してください。"
            self.status_message.set(message)
            messagebox.showwarning("クリップボード読み込み", message)
            return

        try:
            df = self.read_clipboard_text(text)
            if df is None or df.empty:
                raise ValueError("有効なデータがありません")

            if self.load_mode.get() == "append":
                was_empty = not self.raw_dfs
                source_id, label = self.get_next_clipboard_source()
                self.file_paths.append(source_id)
                self.raw_dfs.append(df)
                self.loaded_file_labels.append(label)
                self.source_default_styles[source_id] = "散布図"
                reset_annotations = was_empty
                status_action = "追加"
            else:
                self.file_paths = [self.clipboard_source_id]
                self.raw_dfs = [df]
                self.loaded_file_labels = ["クリップボード"]
                self.source_default_styles = {self.clipboard_source_id: "散布図"}
                reset_annotations = True
                status_action = "読み込み"

            self.update_file_settings_ui()

            self.update_column_info(reset_annotations=reset_annotations)
            self.status_message.set(f"クリップボードから{status_action}完了 ({len(df)}行, {len(df.columns)}列)")
            self.schedule_update()
        except Exception as e:
            message = f"クリップボード読み込みエラー: {str(e)}"
            self.status_message.set(message)
            messagebox.showerror("クリップボード読み込み", message)

    def append_file_data(self, filenames):
        self.status_message.set("データを追加読み込み中...")
        self.root.update()

        errors = []
        skipped = []
        added_count = 0

        for path in filenames:
            if path in self.file_paths:
                skipped.append(os.path.basename(path))
                continue

            try:
                df = self.read_data_file(path)
                if df is not None and not df.empty:
                    self.file_paths.append(path)
                    self.raw_dfs.append(df)
                    self.loaded_file_labels.append(os.path.splitext(os.path.basename(path))[0])
                    added_count += 1
                else:
                    errors.append(f"{os.path.basename(path)}: 有効なデータなし")
            except Exception as e:
                errors.append(f"{os.path.basename(path)}: {str(e)}")

        if added_count > 0:
            self.update_file_settings_ui()
            self.update_column_info(reset_annotations=False)
            self.schedule_update()

        if errors:
            self.status_message.set("追加時に一部エラー: " + " / ".join(errors[:2]))
        elif skipped and added_count == 0:
            self.status_message.set("選択したファイルはすでに読み込み済みです")
        elif skipped:
            self.status_message.set(f"追加完了 ({added_count}個、重複{len(skipped)}個を除外)")
        else:
            self.status_message.set(f"追加完了 ({added_count}個、合計{len(self.raw_dfs)}個)")

    def get_next_clipboard_source(self):
        if self.clipboard_source_id not in self.file_paths:
            return self.clipboard_source_id, "クリップボード"

        index = 2
        while f"{self.clipboard_source_id}{index}" in self.file_paths:
            index += 1
        return f"{self.clipboard_source_id}{index}", f"クリップボード{index}"

    def read_clipboard_text(self, text):
        text = text.replace("\r\n", "\n").replace("\r", "\n").lstrip("\ufeff")
        if not text.strip():
            raise ValueError("クリップボードが空です")

        first_line = next((line for line in text.splitlines() if line.strip()), "")
        delimiter = "\t" if "\t" in first_line else ","
        rows = list(csv.reader(io.StringIO(text), delimiter=delimiter))

        header_input = self.header_row.get().strip()
        if header_input == "":
            header_idx = self.find_header_row_from_rows(rows)
        else:
            try:
                header_idx = int(header_input)
            except ValueError:
                header_idx = None

        df = pd.read_csv(io.StringIO(text), header=header_idx, sep=delimiter, engine="python")
        df = df.dropna(how="all")
        return self._parse_ltspice_ac_data(df)

    def update_file_settings_ui(self):
        """選択されたファイルごとの倍率設定UIを動的に生成する"""
        old_settings = {}
        for path, settings in self.file_settings.items():
            old_settings[path] = (
                settings['x_prefix'].get(),
                settings['y_prefix'].get()
            )

        for widget in self.file_settings_container.winfo_children():
            widget.destroy()
            
        self.file_settings.clear()
        
        if not self.file_paths:
            self.update_preview_selector()
            return
            
        ttk.Label(self.file_settings_container, text="ファイル別 軸の接頭辞設定 (例: 秒→ns は X=n、V→mV は Y=m)", font=("", 9, "bold")).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 5))
        ttk.Label(self.file_settings_container, text="X接頭辞").grid(row=1, column=1, padx=2)
        ttk.Label(self.file_settings_container, text="Y接頭辞").grid(row=1, column=2, padx=2)
        ttk.Label(self.file_settings_container, text="操作").grid(row=1, column=3, padx=2)
        
        for i, path in enumerate(self.file_paths):
            filename = self.get_source_display_name(i)
            display_name = filename if len(filename) <= 18 else f"{filename[:15]}..."
            
            old_x, old_y = old_settings.get(path, ("", ""))
            x_var = tk.StringVar(value=old_x)
            y_var = tk.StringVar(value=old_y)
            x_var.trace_add("write", self.schedule_update)
            y_var.trace_add("write", self.schedule_update)
            
            self.file_settings[path] = {'x_prefix': x_var, 'y_prefix': y_var}
            
            ttk.Label(self.file_settings_container, text=f"{i+1}. {display_name}", width=20).grid(row=i+2, column=0, sticky="w", padx=2, pady=2)
            ttk.Combobox(self.file_settings_container, textvariable=x_var, values=self.prefixes, width=5, state="readonly").grid(row=i+2, column=1, padx=2)
            ttk.Combobox(self.file_settings_container, textvariable=y_var, values=self.prefixes, width=5, state="readonly").grid(row=i+2, column=2, padx=2)
            ttk.Button(self.file_settings_container, text="削除", command=lambda idx=i: self.remove_loaded_data(idx), width=6).grid(row=i+2, column=3, padx=2)

        self.update_preview_selector()

    def get_source_display_name(self, idx):
        if idx < 0 or idx >= len(self.file_paths):
            return ""
        path = self.file_paths[idx]
        if str(path).startswith(self.clipboard_source_id):
            return self.loaded_file_labels[idx] if idx < len(self.loaded_file_labels) else "クリップボード"
        return os.path.basename(path)

    def remove_loaded_data(self, idx):
        if idx < 0 or idx >= len(self.file_paths):
            return

        removed_path = self.file_paths.pop(idx)
        if idx < len(self.raw_dfs):
            self.raw_dfs.pop(idx)
        if idx < len(self.loaded_file_labels):
            removed_label = self.loaded_file_labels.pop(idx)
        else:
            removed_label = str(removed_path)

        self.file_settings.pop(removed_path, None)
        self.source_default_styles.pop(removed_path, None)
        self.anno_click_positions = [None for _ in self.anno_click_positions]
        self.update_file_settings_ui()
        self.update_column_info(reset_annotations=True)
        self.status_message.set(f"削除しました: {removed_label}")
        self.schedule_update()

    def handle_ctrl_v(self, event):
        widget_class = event.widget.winfo_class() if event and event.widget else ""
        if widget_class in ("Entry", "TEntry", "Text", "TCombobox", "Spinbox", "TSpinbox"):
            return
        self.load_clipboard_data()
        return "break"

    def update_preview_selector(self):
        if not hasattr(self, "preview_source_combo"):
            return

        values = [f"{i+1}. {self.get_source_display_name(i)}" for i in range(len(self.file_paths))]
        self.preview_source_combo.configure(values=values)

        if not values:
            self.preview_source_var.set("")
            self.update_data_preview()
            return

        if self.preview_source_var.get() not in values:
            self.preview_source_var.set(values[0])
        else:
            self.update_data_preview()

    def update_data_preview(self):
        if not hasattr(self, "preview_tree"):
            return

        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        self.preview_tree["columns"] = []

        if not self.raw_dfs:
            return

        selected = self.preview_source_var.get()
        try:
            idx = int(selected.split(".", 1)[0]) - 1 if selected else 0
        except ValueError:
            idx = 0

        if idx < 0 or idx >= len(self.raw_dfs):
            return

        df = self.raw_dfs[idx].head(10)
        columns = [str(col) for col in df.columns[:10]]
        self.preview_tree["columns"] = columns

        for col in columns:
            self.preview_tree.heading(col, text=col)
            self.preview_tree.column(col, width=90, stretch=False)

        for _, row in df.iloc[:, :10].iterrows():
            values = []
            for val in row:
                if pd.isna(val):
                    values.append("")
                else:
                    values.append(str(val)[:24])
            self.preview_tree.insert("", tk.END, values=values)

    def export_graph(self):
        if not self.raw_dfs:
            messagebox.showwarning("出力", "出力するグラフがありません。")
            return

        filename = filedialog.asksaveasfilename(
            title="グラフを出力",
            defaultextension=".png",
            filetypes=[
                ("PNG画像", "*.png"),
                ("PDF", "*.pdf"),
                ("SVG", "*.svg"),
                ("すべてのファイル", "*.*"),
            ]
        )
        if not filename:
            return

        try:
            self.fig.savefig(filename, dpi=300, bbox_inches="tight")
            self.status_message.set(f"出力しました: {filename}")
        except Exception as e:
            messagebox.showerror("出力エラー", str(e))

    def collect_series_settings(self):
        settings = []
        for i, key in enumerate(self.series_keys):
            settings.append({
                "key": key,
                "show": self.show_series_vars[i].get() if i < len(self.show_series_vars) else True,
                "text": self.anno_text_vars[i].get() if i < len(self.anno_text_vars) else "",
                "style": self.anno_style_vars[i].get() if i < len(self.anno_style_vars) else "折れ線",
                "mode": self.anno_mode_vars[i].get() if i < len(self.anno_mode_vars) else "自動",
                "position": self.anno_click_positions[i] if i < len(self.anno_click_positions) else None,
                "color": self.series_color_vars[i].get() if i < len(self.series_color_vars) else "自動",
                "marker": self.marker_style_vars[i].get() if i < len(self.marker_style_vars) else "o",
                "line": self.line_style_vars[i].get() if i < len(self.line_style_vars) else "実線",
            })
        return settings

    def save_settings(self):
        filename = filedialog.asksaveasfilename(
            title="設定を保存",
            defaultextension=".json",
            filetypes=[("JSON", "*.json"), ("すべてのファイル", "*.*")]
        )
        if not filename:
            return

        sources = []
        for i, path in enumerate(self.file_paths):
            source = {
                "path": path,
                "label": self.loaded_file_labels[i] if i < len(self.loaded_file_labels) else self.get_source_display_name(i),
                "default_style": self.source_default_styles.get(path, "折れ線"),
                "x_prefix": self.file_settings[path]['x_prefix'].get() if path in self.file_settings else "",
                "y_prefix": self.file_settings[path]['y_prefix'].get() if path in self.file_settings else "",
            }
            if str(path).startswith(self.clipboard_source_id) and i < len(self.raw_dfs):
                source["clipboard_csv"] = self.raw_dfs[i].to_csv(index=False)
            sources.append(source)

        data = {
            "version": 1,
            "sources": sources,
            "ui": {
                "load_mode": self.load_mode.get(),
                "header_row": self.header_row.get(),
                "x_col_index": self.x_col_index.get(),
                "x_min": self.x_min.get(),
                "x_max": self.x_max.get(),
                "y_min": self.y_min.get(),
                "y_max": self.y_max.get(),
                "x_label": self.x_label.get(),
                "y_label": self.y_label.get(),
                "show_y_minmax": self.show_y_minmax.get(),
                "show_annotations": self.show_annotations.get(),
                "x_log_scale": self.x_log_scale.get(),
                "base_font_size": self.base_font_size.get(),
                "line_width": self.line_width.get(),
                "marker_size": self.marker_size.get(),
            },
            "series": self.collect_series_settings(),
        }

        try:
            with open(filename, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self.status_message.set(f"設定を保存しました: {filename}")
        except Exception as e:
            messagebox.showerror("設定保存エラー", str(e))

    def load_settings(self):
        filename = filedialog.askopenfilename(
            title="設定を読み込み",
            filetypes=[("JSON", "*.json"), ("すべてのファイル", "*.*")]
        )
        if not filename:
            return

        try:
            with open(filename, "r", encoding="utf-8") as f:
                data = json.load(f)

            ui = data.get("ui", {})
            self.load_mode.set(ui.get("load_mode", "replace"))
            self.header_row.set(ui.get("header_row", ""))
            self.x_col_index.set(int(ui.get("x_col_index", 0)))
            self.x_min.set(ui.get("x_min", ""))
            self.x_max.set(ui.get("x_max", ""))
            self.y_min.set(ui.get("y_min", ""))
            self.y_max.set(ui.get("y_max", ""))
            self.x_label.set(ui.get("x_label", "x軸"))
            self.y_label.set(ui.get("y_label", "y軸"))
            self.show_y_minmax.set(bool(ui.get("show_y_minmax", False)))
            self.show_annotations.set(bool(ui.get("show_annotations", True)))
            self.x_log_scale.set(bool(ui.get("x_log_scale", False)))
            self.base_font_size.set(int(ui.get("base_font_size", 14)))
            self.line_width.set(float(ui.get("line_width", 2.0)))
            self.marker_size.set(int(ui.get("marker_size", 40)))

            self.file_paths = []
            self.raw_dfs = []
            self.loaded_file_labels = []
            self.source_default_styles = {}

            source_prefixes = {}
            errors = []
            for source in data.get("sources", []):
                path = source.get("path", "")
                label = source.get("label") or os.path.basename(path) or "データ"
                try:
                    if str(path).startswith(self.clipboard_source_id) and source.get("clipboard_csv"):
                        df = pd.read_csv(io.StringIO(source["clipboard_csv"]))
                        df = self._parse_ltspice_ac_data(df)
                    elif path and os.path.exists(path):
                        df = self.read_data_file(path)
                    else:
                        errors.append(f"{label}: 元データが見つかりません")
                        continue

                    self.file_paths.append(path)
                    self.raw_dfs.append(df)
                    self.loaded_file_labels.append(label)
                    self.source_default_styles[path] = source.get("default_style", "折れ線")
                    source_prefixes[path] = (source.get("x_prefix", ""), source.get("y_prefix", ""))
                except Exception as e:
                    errors.append(f"{label}: {str(e)}")

            self.update_file_settings_ui()
            for path, (x_prefix, y_prefix) in source_prefixes.items():
                if path in self.file_settings:
                    self.file_settings[path]['x_prefix'].set(x_prefix)
                    self.file_settings[path]['y_prefix'].set(y_prefix)

            self.update_column_info(reset_annotations=True)
            self.apply_series_settings(data.get("series", []))
            self.update_data_preview()
            self.schedule_update()

            if errors:
                self.status_message.set("設定読み込み完了（一部エラーあり）: " + " / ".join(errors[:2]))
            else:
                self.status_message.set(f"設定を読み込みました: {filename}")
        except Exception as e:
            messagebox.showerror("設定読み込みエラー", str(e))

    def apply_series_settings(self, saved_series):
        by_key = {item.get("key"): item for item in saved_series if item.get("key")}
        for i, key in enumerate(self.series_keys):
            item = by_key.get(key)
            if item is None and i < len(saved_series):
                item = saved_series[i]
            if not item:
                continue

            self.show_series_vars[i].set(bool(item.get("show", True)))
            self.anno_text_vars[i].set(item.get("text", self.anno_text_vars[i].get()))
            self.anno_style_vars[i].set(item.get("style", self.anno_style_vars[i].get()))
            self.anno_mode_vars[i].set(item.get("mode", "自動"))
            pos = item.get("position")
            self.anno_click_positions[i] = tuple(pos) if isinstance(pos, list) and len(pos) == 2 else pos
            self.series_color_vars[i].set(item.get("color", "自動"))
            self.marker_style_vars[i].set(item.get("marker", "o"))
            self.line_style_vars[i].set(item.get("line", "実線"))
            self.update_annotation_status(i)

    def load_all_data(self):
        self.raw_dfs = []
        self.loaded_file_labels = []
        loaded_paths = []
        errors = []
        
        self.status_message.set("データを読み込み中...")
        self.root.update()

        for path in list(self.file_paths):
            try:
                df = self.read_data_file(path)
                if df is not None and not df.empty:
                    loaded_paths.append(path)
                    self.raw_dfs.append(df)
                    self.loaded_file_labels.append(os.path.splitext(os.path.basename(path))[0])
                else:
                    errors.append(f"{os.path.basename(path)}: 有効なデータなし")
            except Exception as e:
                errors.append(f"{os.path.basename(path)}: {str(e)}")

        self.file_paths = loaded_paths
        self.update_file_settings_ui()
                
        if errors:
            self.status_message.set("一部エラー: " + " / ".join(errors[:2]))
        else:
            self.status_message.set(f"読み込み完了 ({len(self.raw_dfs)}個のファイル)")
            
        self.update_column_info()
        self.schedule_update()

    def get_multiplier(self, prefix):
        # データ値を「選んだ接頭辞つき単位」で表示するための倍率。
        # 例: 秒データをns表示にする場合は 1 s = 1e9 ns なので、n は 1e9 倍。
        mapping = {
            "": 1,
            "p": 1e12,
            "n": 1e9,
            "μ": 1e6,
            "µ": 1e6,
            "u": 1e6,
            "m": 1e3,
            "k": 1e-3,
            "M": 1e-6,
            "G": 1e-9,
        }
        return mapping.get(prefix, 1)

    def read_data_file(self, file_path):
        header_input = self.header_row.get().strip()
        if header_input == "":
            header_idx = self.find_header_row(file_path)
        else:
            try:
                header_idx = int(header_input)
            except ValueError:
                header_idx = None
            
        if file_path.lower().endswith(('.xls', '.xlsx')):
            try:
                df = pd.read_excel(file_path, header=header_idx)
            except ImportError:
                raise ImportError("Excelを読み込むには 'openpyxl' が必要です。(pip install openpyxl)")
        else:
            encodings = ['utf-8', 'cp932', 'shift_jis', 'latin-1']
            df = None
            for enc in encodings:
                try:
                    # タブ区切り自動判定
                    with open(file_path, 'r', encoding=enc) as f:
                        first_line = f.readline()
                        delimiter = '\t' if '\t' in first_line else ','
                    df = pd.read_csv(file_path, header=header_idx, encoding=enc, sep=delimiter, engine='python')
                    break
                except UnicodeDecodeError:
                    continue
            if df is None:
                raise ValueError("対応する文字コードが見つかりません")
        
        return self._parse_ltspice_ac_data(df)

    def find_header_row(self, file_path):
        if file_path.lower().endswith(('.xls', '.xlsx')):
            try:
                df_temp = pd.read_excel(file_path, header=None, nrows=100)
                rows = df_temp.values.tolist()
            except Exception:
                return None
        else:
            encodings = ['utf-8', 'cp932', 'shift_jis', 'latin-1']
            lines = []
            for enc in encodings:
                try:
                    with open(file_path, 'r', encoding=enc) as f:
                        lines = [f.readline() for _ in range(100)]
                    break
                except UnicodeDecodeError:
                    pass
            if not lines: return None
            delimiter = '\t' if '\t' in lines[0] else ','
            reader = csv.reader([line for line in lines if line], delimiter=delimiter)
            rows = list(reader)

        return self.find_header_row_from_rows(rows)

    def find_header_row_from_rows(self, rows):
        data_start_idx = -1
        # LTspiceの (XX dB, YY °) 形式を許容する正規表現（堅牢化）
        ltspice_pattern = re.compile(r'^\s*\(\s*[+-]?\d+.*\s*dB\s*,\s*[+-]?\d+.*$')

        for i, row in enumerate(rows):
            if len(row) < 2: continue
            if not any(str(cell).strip() for cell in row if pd.notna(cell)): continue

            is_numeric = True
            numeric_count = 0
            for val in row:
                if pd.isna(val) or not str(val).strip(): continue
                val_str = str(val).strip()
                if ltspice_pattern.match(val_str):
                    numeric_count += 1
                    continue
                try:
                    float(val_str)
                    numeric_count += 1
                except ValueError:
                    is_numeric = False
                    break
            
            if is_numeric and numeric_count > 0:
                data_start_idx = i
                break

        if data_start_idx == -1: return None
        
        candidate_idx = None
        for i in range(data_start_idx - 1, -1, -1):
            if any(str(cell).strip() for cell in rows[i] if pd.notna(cell)):
                candidate_idx = i
                break
        return candidate_idx

    def _parse_ltspice_ac_data(self, df):
        """LTspiceの (20dB, -45°) のようなデータを2つの数値列に分解する"""
        pattern = re.compile(r'^\s*\(\s*([+-]?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)\s*dB\s*,\s*([+-]?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)\s*.*?$')
        new_cols = {}
        cols_to_drop = []
        
        for col in df.columns:
            if df[col].dtype == object:
                valid_val = df[col].dropna().astype(str).str.strip().loc[lambda x: x != ""]
                # 最初の値が(XXdB にマッチするかチェック
                if len(valid_val) > 0 and re.match(r'^\s*\(\s*[+-]?\d+.*\s*dB', valid_val.iloc[0]):
                    db_vals = []
                    deg_vals = []
                    for val in df[col]:
                        if pd.isna(val):
                            db_vals.append(np.nan); deg_vals.append(np.nan)
                            continue
                        m = pattern.match(str(val))
                        if m:
                            db_vals.append(float(m.group(1)))
                            deg_vals.append(float(m.group(2)))
                        else:
                            db_vals.append(np.nan); deg_vals.append(np.nan)
                    
                    new_cols[f"{col} (dB)"] = db_vals
                    new_cols[f"{col} (°)"] = deg_vals
                    cols_to_drop.append(col)
                    
        if new_cols:
            df = df.drop(columns=cols_to_drop)
            for k, v in new_cols.items():
                df[k] = v
        return df

    def build_series_metadata(self):
        """現在読み込まれているデータから、プロット可能な系列候補を作る"""
        try:
            x_idx = self.x_col_index.get()
        except Exception:
            x_idx = 0

        series = []
        multiple_files = len(self.raw_dfs) > 1

        for file_i, df in enumerate(self.raw_dfs):
            if x_idx >= len(df.columns):
                continue

            file_path = self.file_paths[file_i]
            file_label = self.loaded_file_labels[file_i] if file_i < len(self.loaded_file_labels) else self.get_source_display_name(file_i)
            x_col_name = df.columns[x_idx]
            default_style = self.source_default_styles.get(file_path, "折れ線")

            for col_i, col in enumerate(df.columns):
                if col_i == x_idx:
                    continue
                default_label = f"{file_label}: {col}" if multiple_files else str(col)
                series.append({
                    'key': f"{file_path}::{col_i}::{col}",
                    'file_i': file_i,
                    'file_path': file_path,
                    'file_label': file_label,
                    'x_idx': x_idx,
                    'col_i': col_i,
                    'col_name': col,
                    'default_label': default_label,
                    'default_style': default_style,
                })
        return series

    def ensure_series_var_count(self, count):
        while len(self.show_series_vars) < count:
            idx = len(self.show_series_vars)
            show_var = tk.BooleanVar(value=True)
            col_var = tk.StringVar(value="-")
            text_var = tk.StringVar()
            style_var = tk.StringVar(value="折れ線")
            mode_var = tk.StringVar(value="自動")
            status_var = tk.StringVar(value="")
            color_var = tk.StringVar(value="自動")
            marker_var = tk.StringVar(value="なし")
            line_var = tk.StringVar(value="実線")

            for var in [show_var, text_var, style_var, mode_var, color_var, marker_var, line_var]:
                var.trace_add("write", self.schedule_update)
            mode_var.trace_add("write", lambda *args, idx=idx: self.update_annotation_status(idx))

            self.show_series_vars.append(show_var)
            self.col_name_vars.append(col_var)
            self.anno_text_vars.append(text_var)
            self.anno_style_vars.append(style_var)
            self.anno_mode_vars.append(mode_var)
            self.anno_status_vars.append(status_var)
            self.anno_click_positions.append(None)
            self.series_color_vars.append(color_var)
            self.marker_style_vars.append(marker_var)
            self.line_style_vars.append(line_var)

        if len(self.show_series_vars) > count:
            self.show_series_vars = self.show_series_vars[:count]
            self.col_name_vars = self.col_name_vars[:count]
            self.anno_text_vars = self.anno_text_vars[:count]
            self.anno_style_vars = self.anno_style_vars[:count]
            self.anno_mode_vars = self.anno_mode_vars[:count]
            self.anno_status_vars = self.anno_status_vars[:count]
            self.anno_click_positions = self.anno_click_positions[:count]
            self.series_color_vars = self.series_color_vars[:count]
            self.marker_style_vars = self.marker_style_vars[:count]
            self.line_style_vars = self.line_style_vars[:count]

    def update_column_info(self, reset_annotations=True):
        """読み込んだデータの系列一覧をUIに反映"""
        try:
            series = self.build_series_metadata()
            self.series_keys = [info['key'] for info in series]
            self.ensure_series_var_count(len(series))

            for i, info in enumerate(series):
                label = info['default_label']
                self.col_name_vars[i].set(label)
                had_text = bool(self.anno_text_vars[i].get())

                if reset_annotations:
                    self.show_series_vars[i].set(True)
                    self.anno_text_vars[i].set(label)
                    self.anno_style_vars[i].set(info['default_style'])
                    self.anno_mode_vars[i].set("自動")
                    self.anno_click_positions[i] = None
                    self.series_color_vars[i].set("自動")
                    self.marker_style_vars[i].set("なし")
                    self.line_style_vars[i].set("実線")
                elif not had_text:
                    self.anno_text_vars[i].set(label)
                    self.anno_style_vars[i].set(info['default_style'])

                self.update_annotation_status(i)

            self.rebuild_series_settings_ui()
        except Exception:
            pass

    def rebuild_series_settings_ui(self):
        for widget in self.anno_rows_container.winfo_children():
            widget.destroy()

        if not self.series_keys:
            ttk.Label(self.anno_rows_container, text="データを読み込むと系列設定が表示されます", foreground="gray").pack(anchor="w")
            return

        for i in range(len(self.series_keys)):
            row = ttk.Frame(self.anno_rows_container)
            row.pack(fill=tk.X, pady=(0, 8))

            top = ttk.Frame(row)
            top.pack(fill=tk.X)
            ttk.Label(top, text=f"{i+1}.", width=3).pack(side=tk.LEFT)
            ttk.Checkbutton(top, variable=self.show_series_vars[i]).pack(side=tk.LEFT)
            ttk.Label(top, textvariable=self.col_name_vars[i], width=28).pack(side=tk.LEFT, padx=3)

            mid = ttk.Frame(row)
            mid.pack(fill=tk.X, pady=(2, 0))
            ttk.Label(mid, text="文字").pack(side=tk.LEFT)
            ttk.Entry(mid, textvariable=self.anno_text_vars[i], width=14).pack(side=tk.LEFT, padx=3)
            ttk.Label(mid, text="種類").pack(side=tk.LEFT)
            ttk.Combobox(mid, textvariable=self.anno_style_vars[i], values=["折れ線", "散布図", "線＋点"], width=7, state="readonly").pack(side=tk.LEFT, padx=3)
            ttk.Label(mid, text="色").pack(side=tk.LEFT)
            ttk.Combobox(mid, textvariable=self.series_color_vars[i], values=self.series_color_palette, width=7).pack(side=tk.LEFT, padx=3)
            ttk.Button(mid, text="選択", command=lambda idx=i: self.choose_series_color(idx), width=5).pack(side=tk.LEFT)

            bottom = ttk.Frame(row)
            bottom.pack(fill=tk.X, pady=(2, 0))
            ttk.Label(bottom, text="点").pack(side=tk.LEFT)
            ttk.Combobox(bottom, textvariable=self.marker_style_vars[i], values=self.marker_options, width=5, state="readonly").pack(side=tk.LEFT, padx=3)
            ttk.Label(bottom, text="線").pack(side=tk.LEFT)
            ttk.Combobox(bottom, textvariable=self.line_style_vars[i], values=self.line_style_options, width=7, state="readonly").pack(side=tk.LEFT, padx=3)
            ttk.Label(bottom, text="配置").pack(side=tk.LEFT)
            ttk.Combobox(bottom, textvariable=self.anno_mode_vars[i], values=["自動", "クリック指定", "固定"], width=8, state="readonly").pack(side=tk.LEFT, padx=3)
            ttk.Button(bottom, text="自動に戻す", command=lambda idx=i: self.reset_annotation_position(idx), width=10).pack(side=tk.LEFT, padx=3)

            ttk.Label(row, textvariable=self.anno_status_vars[i], foreground="blue", font=("", 8), wraplength=390).pack(anchor="w", padx=28)

    def choose_series_color(self, idx):
        if idx >= len(self.series_color_vars):
            return
        color = colorchooser.askcolor(title="系列色を選択")
        if color and color[1]:
            self.series_color_vars[idx].set(color[1])

    def get_marker_style(self, idx):
        marker = self.marker_style_vars[idx].get() if idx < len(self.marker_style_vars) else "o"
        return None if marker == "なし" else marker

    def get_line_style(self, idx):
        line = self.line_style_vars[idx].get() if idx < len(self.line_style_vars) else "実線"
        mapping = {"実線": "-", "破線": "--", "点線": ":", "一点鎖線": "-.", "なし": "None"}
        return mapping.get(line, "-")

    def update_annotation_status(self, idx):
        if idx >= len(self.anno_status_vars): return
        mode = self.anno_mode_vars[idx].get()
        pos = self.anno_click_positions[idx]

        if mode == "自動":
            self.anno_status_vars[idx].set("")
        elif mode == "クリック指定":
            self.anno_status_vars[idx].set("※ グラフ上をクリック (指定後、自動で「固定」になります)")
        elif mode == "固定":
            if pos is None:
                self.anno_status_vars[idx].set("※ 位置がありません。一度「クリック指定」を選んでください。")
            else:
                self.anno_status_vars[idx].set(f"固定済 (x={pos[0]:.2g}, y={pos[1]:.2g})")

    def reset_annotation_position(self, idx):
        self.anno_click_positions[idx] = None
        self.anno_mode_vars[idx].set("自動")
        self.schedule_update()

    def update_plot(self):
        """現在の設定に基づいてグラフを再描画する"""
        if not self.raw_dfs:
            self.ax.clear()
            self.ax.text(0.5, 0.5, "データを読み込むとここにグラフが表示されます", 
                         horizontalalignment='center', verticalalignment='center', 
                         transform=self.ax.transAxes, color='gray')
            self.canvas.draw()
            return
            
        try:
            series_infos = []
            
            # 各データ系列を準備
            for ui_index, meta in enumerate(self.build_series_metadata()):
                df = self.raw_dfs[meta['file_i']]
                file_path = meta['file_path']
                x_idx = meta['x_idx']
                
                # ファイル個別の接頭辞変換倍率を取得
                x_mult = 1; y_mult = 1
                if file_path in self.file_settings:
                    x_mult = self.get_multiplier(self.file_settings[file_path]['x_prefix'].get())
                    y_mult = self.get_multiplier(self.file_settings[file_path]['y_prefix'].get())

                if x_idx >= len(df.columns): continue

                x_series = pd.to_numeric(df.iloc[:, x_idx], errors='coerce') * x_mult

                y_series = pd.to_numeric(df.iloc[:, meta['col_i']], errors='coerce') * y_mult
                series_data = pd.DataFrame({'x': x_series, 'y': y_series}).dropna()
                if series_data.empty:
                    continue

                series_infos.append({
                    'index': ui_index,
                    'sx': series_data['x'].values,
                    'sy': series_data['y'].values,
                    'default_label': meta['default_label'],
                    'default_style': meta['default_style'],
                })

            if not series_infos:
                self.status_message.set("プロット可能な数値データがありません")
                return

            self.ax.clear()
            
            # フォントサイズ等のデザイン設定の取得
            try:
                base_fs = int(self.base_font_size.get())
                lw = float(self.line_width.get())
                ms = int(self.marker_size.get())
            except ValueError:
                base_fs = 14; lw = 2.0; ms = 40
            
            # 軸の装飾設定
            self.ax.minorticks_on()
            self.ax.tick_params(axis='both', which='major', labelsize=base_fs * 0.9)
            self.ax.tick_params(axis='both', which='both', direction='in', top=True, right=True)
            self.ax.yaxis.set_minor_locator(AutoMinorLocator())
            self.ax.grid(True, which='major', linestyle='--', alpha=0.7)
            self.ax.grid(True, which='minor', linestyle=':', alpha=0.4)
            
            if self.x_log_scale.get():
                self.ax.set_xscale('log')
            else:
                self.ax.set_xscale('linear')
                self.ax.xaxis.set_minor_locator(AutoMinorLocator())

            colors = plt.cm.tab10.colors 
            color_cycle = itertools.cycle(colors)
            all_lines_data = []

            # プロット実行
            for info in series_infos:
                i = info['index']
                
                # チェックボックスによる表示/非表示の判定
                if i < len(self.show_series_vars) and not self.show_series_vars[i].get():
                    info['is_visible'] = False
                    continue
                info['is_visible'] = True

                auto_color = next(color_cycle)
                selected_color = self.series_color_vars[i].get() if i < len(self.series_color_vars) else "自動"
                c = auto_color if selected_color in ("", "自動") else selected_color
                info['color'] = c
                
                # スタイルとラベルの適用
                display_text = info['default_label']
                style = info.get('default_style', "折れ線")
                if i < len(self.anno_text_vars):
                    if self.anno_text_vars[i].get(): display_text = self.anno_text_vars[i].get()
                    style = self.anno_style_vars[i].get()
                
                info['display_text'] = self.format_custom_text(display_text)

                sx, sy = info['sx'], info['sy']
                marker = self.get_marker_style(i)
                line_style = self.get_line_style(i)
                plot_marker_size = max(2.0, ms ** 0.5)

                if style == "折れ線":
                    self.ax.plot(sx, sy, color=c, linestyle=line_style, marker=marker, markersize=plot_marker_size if marker else 0,
                                 linewidth=lw, zorder=3, label=info['display_text'])
                elif style == "散布図":
                    self.ax.scatter(sx, sy, color=c, marker=marker or "o", s=ms, zorder=4, label=info['display_text'])
                else: # 線＋点
                    self.ax.plot(sx, sy, color=c, linestyle=line_style, marker=marker or "o", markersize=plot_marker_size,
                                 linewidth=lw*0.5, zorder=4, label=info['display_text'])
                    
                # スマート配置用のデータ縮引き（重すぎないようにする）
                if len(sx) > 500:
                    indices = np.linspace(0, len(sx)-1, 500).astype(int)
                    all_lines_data.append((sx[indices], sy[indices]))
                else:
                    all_lines_data.append((sx, sy))

            # 手動での範囲指定があれば適用（空欄の場合はMatplotlibの自動調整に任せる）
            try:
                x_min_str, x_max_str = self.x_min.get().strip(), self.x_max.get().strip()
                if x_min_str or x_max_str:
                    x_min_val = float(x_min_str) if x_min_str else None
                    x_max_val = float(x_max_str) if x_max_str else None
                    self.ax.set_xlim(left=x_min_val, right=x_max_val)
            except ValueError:
                pass
                
            try:
                y_min_str, y_max_str = self.y_min.get().strip(), self.y_max.get().strip()
                if y_min_str or y_max_str:
                    y_min_val = float(y_min_str) if y_min_str else None
                    y_max_val = float(y_max_str) if y_max_str else None
                    self.ax.set_ylim(bottom=y_min_val, top=y_max_val)
            except ValueError:
                pass

            # 描画範囲の確定値を取得
            view_limits = self.ax.get_xlim() + self.ax.get_ylim()

            # アノテーションの描画
            self.draw_annotations(series_infos, all_lines_data, view_limits, base_fs)

            self.ax.set_xlabel(self.format_custom_text(self.x_label.get()), fontsize=base_fs)
            self.ax.set_ylabel(self.format_custom_text(self.y_label.get()), fontsize=base_fs)
            self.fig.tight_layout()
            self.canvas.draw()
            self.status_message.set(f"プレビュー更新完了 ({len(series_infos)} 系列)")

        except Exception as e:
            self.status_message.set(f"描画エラー: {str(e)}")

    def format_custom_text(self, text):
        """
        文字列中の特定パターンをMatplotlibのMathTextに変換する
        !...! -> 斜体
        _{...} -> 下付き
        ^{...} -> 上付き
        """
        if not isinstance(text, str):
            return text
            
        def repl(match):
            m1 = match.group(1) # !...!
            m2 = match.group(2) # _{...}
            m3 = match.group(3) # ^{...}
            
            if m1 is not None:
                # MathText内でスペースが無視されるのを防ぐ
                content = m1[1:-1].replace(' ', r'\ ')
                return f'$\\mathit{{{content}}}$'
            elif m2 is not None:
                content = m2[2:-1].replace(' ', r'\ ')
                return f'$_{{{content}}}$'
            elif m3 is not None:
                content = m3[2:-1].replace(' ', r'\ ')
                return f'$^{{{content}}}$'
            return match.group(0)
            
        return re.sub(r'(!.*?!)|(_\{.*?\})|(\^\{.*?\})', repl, text)

    def draw_annotations(self, series_infos, all_lines_data, view_limits, base_fs=14):
        placed_bboxes = []
        final_x_min, final_x_max, final_y_min, final_y_max = view_limits
        is_log = self.x_log_scale.get()

        for info in series_infos:
            if not info.get('is_visible', True):
                continue
                
            i = info['index']
            sx, sy, c, display_text = info['sx'], info['sy'], info['color'], info['display_text']

            if self.show_annotations.get():
                mode = self.anno_mode_vars[i].get() if i < len(self.anno_mode_vars) else '自動'
                pos = self.anno_click_positions[i] if i < len(self.anno_click_positions) else None

                if mode in ['クリック指定', '固定'] and pos is not None:
                    text_x, text_y = pos
                    # 対数スケール時の距離計算の補正
                    if is_log:
                        mask = (sx > 0) & (sx >= final_x_min) & (sx <= final_x_max)
                        vis_sx = sx[mask] if np.any(mask) else sx[sx > 0]
                        vis_sy = sy[mask] if np.any(mask) else sy[sx > 0]
                        if len(vis_sx) > 0:
                            nx = (np.log10(vis_sx) - np.log10(text_x)) / max(1e-10, np.log10(final_x_max) - np.log10(final_x_min))
                            ny = (vis_sy - text_y) / max(1e-10, final_y_max - final_y_min)
                            min_idx = np.argmin(nx**2 + ny**2)
                            target_x, target_y = vis_sx[min_idx], vis_sy[min_idx]
                        else:
                            target_x, target_y = text_x, text_y
                    else:
                        mask = (sx >= final_x_min) & (sx <= final_x_max)
                        vis_sx = sx[mask] if np.any(mask) else sx
                        vis_sy = sy[mask] if np.any(mask) else sy
                        nx = (vis_sx - text_x) / max(1e-10, final_x_max - final_x_min)
                        ny = (vis_sy - text_y) / max(1e-10, final_y_max - final_y_min)
                        min_idx = np.argmin(nx**2 + ny**2)
                        target_x, target_y = vis_sx[min_idx], vis_sy[min_idx]
                else:
                    (target_x, target_y), (text_x, text_y), bw, bh = self.find_smart_position((sx, sy), all_lines_data, placed_bboxes, view_limits)
                    placed_bboxes.append((text_x, text_y, bw, bh))

                self.ax.annotate(display_text, xy=(target_x, target_y), xytext=(text_x, text_y),
                            arrowprops=dict(arrowstyle="->", color=c, connectionstyle="arc3,rad=0"), 
                            bbox=dict(boxstyle="round,pad=0.3", fc="white", ec=c, lw=max(1.0, base_fs/7), alpha=0.9),
                            fontsize=max(8.0, base_fs * 0.85), color=c, va='center', fontweight='bold', zorder=10)

            if self.show_y_minmax.get():
                self.annotate_y_minmax(sx, sy, display_text, c, view_limits, base_fs)

    def find_smart_position(self, target_series, all_lines, placed_bboxes, limits):
        series_x, series_y = target_series
        xmin, xmax, ymin, ymax = limits
        
        # 簡易的なスマート配置（高速化のため右端付近を優先）
        mask = (series_x >= xmin) & (series_x <= xmax) & (series_y >= ymin) & (series_y <= ymax)
        valid_idx = np.where(mask)[0]
        
        if len(valid_idx) == 0:
            return (series_x[-1], series_y[-1]), (series_x[-1], series_y[-1]), 0, 0
            
        target_idx = valid_idx[int(len(valid_idx) * 0.8)]
        tx, ty = series_x[target_idx], series_y[target_idx]
        
        # ログスケールの場合は表示幅を考慮
        offset_x = tx * 1.5 if self.x_log_scale.get() else tx + (xmax - xmin)*0.05
        offset_y = ty + (ymax - ymin)*0.05
        
        return (tx, ty), (offset_x, offset_y), (xmax-xmin)*0.1, (ymax-ymin)*0.1

    def annotate_y_minmax(self, sx, sy, text, color, limits, base_fs=14):
        xmin, xmax, ymin, ymax = limits
        mask = (sx >= xmin) & (sx <= xmax) & (sy >= ymin) & (sy <= ymax)
        if not np.any(mask): return
        
        vis_sx, vis_sy = sx[mask], sy[mask]
        min_idx, max_idx = np.argmin(vis_sy), np.argmax(vis_sy)
        
        points = [("min", vis_sx[min_idx], vis_sy[min_idx], (8, -18), "top")]
        if min_idx != max_idx:
            points.append(("max", vis_sx[max_idx], vis_sy[max_idx], (8, 10), "bottom"))
            
        for label, px, py, offset, va in points:
            self.ax.scatter(px, py, color=color, s=base_fs*2.5, zorder=5)
            self.ax.annotate(f"{text}\n{label} y={py:.4g}", xy=(px, py), xytext=offset, textcoords="offset points",
                             fontsize=max(8.0, base_fs * 0.75), color=color, va=va, bbox=dict(boxstyle="round,pad=0.2", fc="white", ec=color, alpha=0.8), zorder=6)

    def on_canvas_click(self, event):
        """プレビューグラフ上でのクリック位置を記録し、再描画する"""
        if not self.show_annotations.get() or event.inaxes != self.ax or event.xdata is None:
            return
            
        # ToolbarがZoom/Panモードの時は無効化
        if self.toolbar.mode: return

        click_targets = [i for i in range(len(self.anno_mode_vars)) if self.anno_mode_vars[i].get() == 'クリック指定']
        if not click_targets: return

        # 未指定のものがあれば優先、なければクリック位置のX座標に最も近い系列を更新
        target_idx = None
        for idx in click_targets:
            if self.anno_click_positions[idx] is None:
                target_idx = idx
                break
                
        if target_idx is None:
            target_idx = click_targets[0] # 簡易的に最初の指定済みを上書き

        self.anno_click_positions[target_idx] = (float(event.xdata), float(event.ydata))
        # クリック後、誤操作を防ぐために自動的に「固定」モードに移行する
        self.anno_mode_vars[target_idx].set("固定")
        self.update_annotation_status(target_idx)
        self.schedule_update()

if __name__ == "__main__":
    root = tk.Tk()
    app = GraphApp(root)
    root.mainloop()

