import tkinter as tk
from tkinter import filedialog, messagebox

class FileSelectFrame(tk.Frame):
    def __init__(self, master, next_callback):
        super().__init__(master)
        self.master = master
        self.next_callback = next_callback
        self.create_widgets()

    def create_widgets(self):
        self.grid_columnconfigure(1, weight=1)

        # タイトル
        tk.Label(self, text="データ移行ツール - ファイル選択", font=("Arial", 14)).grid(row=0, column=0, columnspan=3, pady=(20,10))

        # 移行元ファイル
        tk.Label(self, text="移行元ファイル:").grid(row=1, column=0, sticky="e", padx=10, pady=5)
        self.src_path = tk.Entry(self)
        self.src_path.grid(row=1, column=1, sticky="ew", padx=5)
        tk.Button(self, text="参照", command=self.select_src_file).grid(row=1, column=2, padx=10)

        # 移行元ヘッダー
        tk.Label(self, text="移行元ヘッダー行番号（例: 1）:").grid(row=2, column=0, sticky="e", padx=10, pady=5)
        self.src_header = tk.Entry(self, width=10)
        self.src_header.grid(row=2, column=1, sticky="w", padx=5)

        # 移行先ファイル
        tk.Label(self, text="移行先ファイル:").grid(row=3, column=0, sticky="e", padx=10, pady=5)
        self.dst_path = tk.Entry(self)
        self.dst_path.grid(row=3, column=1, sticky="ew", padx=5)
        tk.Button(self, text="参照", command=self.select_dst_file).grid(row=3, column=2, padx=10)

        # 移行先ヘッダー
        tk.Label(self, text="移行先ヘッダー行番号（例: 1）:").grid(row=4, column=0, sticky="e", padx=10, pady=5)
        self.dst_header = tk.Entry(self, width=10)
        self.dst_header.grid(row=4, column=1, sticky="w", padx=5)

        # ボタンエリア
        btn_frame = tk.Frame(self)
        btn_frame.grid(row=5, column=0, columnspan=3, pady=20)
        btn_frame.grid_columnconfigure(1, weight=1)

        tk.Button(btn_frame, text="キャンセル", command=self.master.destroy).grid(row=0, column=0, padx=20)
        tk.Button(btn_frame, text="次へ", command=self.on_next).grid(row=0, column=1, padx=20)

    def select_src_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV/Excel files", "*.csv *.xlsx *.xls")])
        if path:
            self.src_path.delete(0, tk.END)
            self.src_path.insert(0, path)

    def select_dst_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV/Excel files", "*.csv *.xlsx *.xls")])
        if path:
            self.dst_path.delete(0, tk.END)
            self.dst_path.insert(0, path)

    def on_next(self):
        if not self.src_path.get() or not self.src_header.get() or not self.dst_path.get() or not self.dst_header.get():
            messagebox.showerror("エラー", "すべての項目を入力してください。")
            return

        try:
            import pandas as pd
            src_df = pd.read_csv(self.src_path.get(), header=int(self.src_header.get()) - 1)
            dst_df = pd.read_csv(self.dst_path.get(), header=int(self.dst_header.get()) - 1)
            src_cols = list(src_df.columns)
            dst_cols = list(dst_df.columns)
        except Exception as e:
            messagebox.showerror("読み込みエラー", str(e))
            return

        self.next_callback(src_df, src_cols, dst_cols)