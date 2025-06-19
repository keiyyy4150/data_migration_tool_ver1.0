import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from logic.exporter import convert_and_export
import pandas as pd

DATA_TYPES = ["str", "int", "float", "bool", "date"]

class MappingFrame(tk.Frame):
    def __init__(self, master, src_df, src_columns, dst_columns, on_back_callback):
        super().__init__(master)
        self.master = master
        self.src_df = src_df
        self.src_columns = src_columns
        self.dst_columns = dst_columns
        self.on_back_callback = on_back_callback
        self.mapping_widgets = []
        self.create_widgets()

    def create_widgets(self):
        tk.Label(self, text="移行カラムマッピング", font=("Arial", 14)).grid(row=0, column=0, columnspan=3, pady=10)

        header_frame = tk.Frame(self)
        header_frame.grid(row=1, column=0, columnspan=3)
        # header
        tk.Label(header_frame, text="移行元カラム", width=20, anchor="w", padx=10).grid(row=0, column=0)
        tk.Label(header_frame, text="移行先カラム", width=30, anchor="w", padx=10).grid(row=0, column=1)
        tk.Label(header_frame, text="データ型", width=10, anchor="w", padx=10).grid(row=0, column=2)

        self.mapping_frame = tk.Frame(self)
        self.mapping_frame.grid(row=2, column=0, columnspan=3)

        for i, src_col in enumerate(self.src_columns):
            tk.Label(self.mapping_frame, text=src_col, width=20, anchor="w").grid(row=i, column=0, padx=5, pady=2)

            # 初期値で一致するものを選択
            initial = src_col if src_col in self.dst_columns else ""
            combo = ttk.Combobox(self.mapping_frame, values=self.dst_columns, state="readonly", width=28)
            combo.set(initial)
            combo.grid(row=i, column=1, padx=5)

            dtype_combo = ttk.Combobox(self.mapping_frame, values=DATA_TYPES, state="readonly", width=10)
            dtype_combo.set("str")
            dtype_combo.grid(row=i, column=2)

            self.mapping_widgets.append((src_col, combo, dtype_combo))

            # ボタンフレーム調整
            btn_frame = tk.Frame(self)
            btn_frame.grid(row=3, column=0, columnspan=3, pady=20)

            tk.Button(btn_frame, text="戻る", width=12, command=self.on_back).pack(side="left", padx=20)
            tk.Button(btn_frame, text="キャンセル", width=12, command=self.master.destroy).pack(side="left", padx=20)
            tk.Button(btn_frame, text="データ移行実行", width=16, command=self.on_execute).pack(side="right", padx=20)

    def on_back(self):
        self.on_back_callback()

    def on_execute(self):
        result = []
        used = set()

        for src, dst_combo, dtype_combo in self.mapping_widgets:
            dst = dst_combo.get()
            dtype = dtype_combo.get()
            if dst == "" or dst in used:
                messagebox.showerror("エラー", f"マッピングが不正です（重複または未選択）。")
                return
            used.add(dst)
            result.append((src, dst, dtype))

        # 出力先のパスを選択
        output_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if not output_path:
            return

        # 実際の出力処理
        try:
            convert_and_export(self.src_df, result, output_path)
            messagebox.showinfo("完了", f"出力が完了しました：\n{output_path}")
        except Exception as e:
            messagebox.showerror("出力エラー", str(e))