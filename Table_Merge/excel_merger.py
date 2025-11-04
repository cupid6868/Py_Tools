import tkinter as tk
from tkinter import filedialog, messagebox, ttk, font
import pandas as pd
import os
import threading
from multiprocessing import Process, Queue
import time
from typing import List, Dict


class ExcelMergerPro:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel智能合并工具 - 多条件匹配版")
        self.root.geometry("1100x700")
        self.root.minsize(1000, 650)

        # 中文字体设置
        self.font = ("SimHei", 10)
        self.small_font = ("SimHei", 9)

        # 数据存储
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.output_path = tk.StringVar(value="合并结果.xlsx")
        self.df1 = None  # 仅预览用
        self.df2 = None  # 仅预览用
        self.selected_cols = []
        self.is_running = False
        self.start_time = 0
        self.total_rows = 0
        self.processed_rows = 0

        # 进程间通信
        self.progress_queue = Queue()
        self.control_queue = Queue()
        self.log_queue = Queue()
        self.merge_process = None

        # 匹配配置（支持多对匹配）
        self.match_pairs: List[Dict] = []

        self.create_widgets()
        self.start_progress_listener()
        self.start_log_listener()

    def create_widgets(self):
        # 主滚动区域（保持不变）
        main_canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=main_canvas.yview)
        scrollable_frame = ttk.Frame(main_canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )

        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)

        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 顶部标题（强调多条件匹配）
        title_font = font.Font(family="SimHei", size=16, weight="bold")
        tk.Label(scrollable_frame, text="Excel智能合并工具（多条件联合匹配）", font=title_font).pack(pady=10)
        tk.Label(scrollable_frame, text="注：所有匹配对必须同时匹配成功才会执行合并", font=("SimHei", 9, "italic"),
                 fg="#666").pack(pady=2)

        # 文件选择区域（保持不变）
        file_frame = tk.Frame(scrollable_frame)
        file_frame.pack(fill=tk.X, padx=20, pady=5)

        # 第一个文件
        tk.Label(file_frame, text="第一个Excel文件：", font=self.font).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        tk.Entry(file_frame, textvariable=self.file1_path, width=60, font=self.font).grid(row=0, column=1, padx=5,
                                                                                          pady=5)
        tk.Button(file_frame, text="浏览", command=self.browse_file1, font=self.font).grid(row=0, column=2, padx=5,
                                                                                           pady=5)
        tk.Button(file_frame, text="预览", command=self.preview_file1, font=self.small_font).grid(row=0, column=3,
                                                                                                  padx=2, pady=5)

        # 第二个文件
        tk.Label(file_frame, text="第二个Excel文件：", font=self.font).grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        tk.Entry(file_frame, textvariable=self.file2_path, width=60, font=self.font).grid(row=1, column=1, padx=5,
                                                                                          pady=5)
        tk.Button(file_frame, text="浏览", command=self.browse_file2, font=self.font).grid(row=1, column=2, padx=5,
                                                                                           pady=5)
        tk.Button(file_frame, text="预览", command=self.preview_file2, font=self.small_font).grid(row=1, column=3,
                                                                                                  padx=2, pady=5)

        # 输出文件
        tk.Label(file_frame, text="输出文件路径：", font=self.font).grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        tk.Entry(file_frame, textvariable=self.output_path, width=60, font=self.font).grid(row=2, column=1, padx=5,
                                                                                           pady=5)
        tk.Button(file_frame, text="选择路径", command=self.browse_output, font=self.font).grid(row=2, column=2, padx=5,
                                                                                                pady=5)

        # 匹配设置区域（强调多条件联合匹配）
        match_frame = tk.LabelFrame(scrollable_frame, text="匹配规则设置（所有条件必须同时满足）", font=self.font)
        match_frame.pack(fill=tk.X, padx=20, pady=10)

        # 匹配对容器
        self.match_pairs_frame = tk.Frame(match_frame)
        self.match_pairs_frame.pack(fill=tk.X, padx=10, pady=5)

        # 添加匹配对按钮
        tk.Button(match_frame, text="添加匹配对（多条件）", command=self.add_match_pair, font=self.font).pack(pady=10)

        # 列选择区域（保持不变）
        col_select_frame = tk.LabelFrame(scrollable_frame, text="选择第二个表中需要合并的列", font=self.font)
        col_select_frame.pack(fill=tk.BOTH, expand=False, padx=20, pady=10)

        self.col_listbox = tk.Listbox(col_select_frame, selectmode=tk.MULTIPLE, font=self.small_font, height=4,
                                      width=100)
        self.col_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10, pady=10)

        col_scrollbar = ttk.Scrollbar(col_select_frame, orient="horizontal", command=self.col_listbox.xview)
        self.col_listbox.configure(xscrollcommand=col_scrollbar.set)
        col_scrollbar.pack(side=tk.BOTTOM, fill=tk.X, padx=10)

        btn_frame = tk.Frame(col_select_frame)
        btn_frame.pack(side=tk.RIGHT, padx=10, pady=10)
        tk.Button(btn_frame, text="全选", command=self.select_all_cols, font=self.small_font).pack(pady=5)
        tk.Button(btn_frame, text="取消全选", command=self.deselect_all_cols, font=self.small_font).pack(pady=5)

        # 日志与预览区域
        log_preview_frame = tk.Frame(scrollable_frame)
        log_preview_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # 左侧文件信息预览
        preview_frame = tk.LabelFrame(log_preview_frame, text="文件信息预览", font=self.font)
        preview_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 左侧文件1信息
        self.frame1 = tk.Frame(preview_frame)
        self.frame1.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=5, pady=5)
        tk.Label(self.frame1, text="文件1信息：", font=self.font).pack(anchor=tk.W)
        self.info1 = tk.Text(self.frame1, height=5, width=45, font=self.small_font, state=tk.DISABLED)
        self.info1.pack(fill=tk.BOTH, expand=True, pady=5)

        # 右侧文件2信息
        self.frame2 = tk.Frame(preview_frame)
        self.frame2.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True, padx=5, pady=5)
        tk.Label(self.frame2, text="文件2信息：", font=self.font).pack(anchor=tk.W)
        self.info2 = tk.Text(self.frame2, height=5, width=45, font=self.small_font, state=tk.DISABLED)
        self.info2.pack(fill=tk.BOTH, expand=True, pady=5)

        # 右侧匹配日志（增强多条件匹配记录）
        log_frame = tk.LabelFrame(log_preview_frame, text="多条件匹配日志（所有条件需同时满足）", font=self.font)
        log_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.log_text = tk.Text(log_frame, height=10, width=45, font=self.small_font, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=5)

        # 进度条、状态和时间统计（保持不变）
        progress_frame = tk.Frame(scrollable_frame)
        progress_frame.pack(fill=tk.X, padx=20, pady=5)

        self.progress = ttk.Progressbar(progress_frame, orient="horizontal", length=100, mode="determinate")
        self.progress.pack(fill=tk.X, side=tk.LEFT, padx=5)

        self.status_var = tk.StringVar(value="就绪")
        tk.Label(progress_frame, textvariable=self.status_var, font=self.font).pack(side=tk.LEFT, padx=10)

        self.time_var = tk.StringVar(value="耗时：--:--")
        tk.Label(progress_frame, textvariable=self.time_var, font=self.font, fg="#666").pack(side=tk.RIGHT, padx=10)

        # 操作按钮（保持不变）
        btn_frame = tk.Frame(scrollable_frame)
        btn_frame.pack(pady=15)

        self.run_btn = tk.Button(btn_frame, text="开始合并", command=self.start_merge_process,
                                 font=("SimHei", 12, "bold"), width=15, bg="#2196F3", fg="white")
        self.run_btn.pack(side=tk.LEFT, padx=10)

        self.cancel_btn = tk.Button(btn_frame, text="取消", command=self.cancel_merge,
                                    font=("SimHei", 12), width=15, bg="#ff9800", fg="white", state=tk.DISABLED)
        self.cancel_btn.pack(side=tk.LEFT, padx=10)

        tk.Button(btn_frame, text="清除选择", command=self.clear_selection,
                  font=("SimHei", 12), width=15).pack(side=tk.LEFT, padx=10)

        tk.Button(btn_frame, text="退出", command=self.root.quit,
                  font=("SimHei", 12), width=15, bg="#f44336", fg="white").pack(side=tk.LEFT, padx=10)

        # 初始添加一对匹配列
        self.add_match_pair()

    # 以下方法（add_match_pair到cancel_merge）与原代码一致，省略重复部分
    # （保持UI交互逻辑不变，仅修改process_merge函数）

    def add_match_pair(self):
        """动态添加匹配对（带独立规则）"""
        pair_frame = tk.Frame(self.match_pairs_frame)
        pair_frame.pack(fill=tk.X, padx=5, pady=3)

        # 表1列选择
        col1_var = tk.StringVar()
        col1_combo = ttk.Combobox(pair_frame, textvariable=col1_var, width=25, font=self.small_font, state="disabled")
        col1_combo.pack(side=tk.LEFT, padx=5)

        # "对"标签
        tk.Label(pair_frame, text="对", font=self.font).pack(side=tk.LEFT, padx=5)

        # 表2列选择
        col2_var = tk.StringVar()
        col2_combo = ttk.Combobox(pair_frame, textvariable=col2_var, width=25, font=self.small_font, state="disabled")
        col2_combo.pack(side=tk.LEFT, padx=5)

        # 匹配规则
        rule_var = tk.StringVar(value="fuzzy")
        rule_frame = tk.Frame(pair_frame)
        tk.Label(rule_frame, text="规则：", font=self.small_font).pack(side=tk.LEFT)
        tk.Radiobutton(rule_frame, text="完全相同", variable=rule_var, value="exact", font=("SimHei", 8)).pack(
            side=tk.LEFT)
        tk.Radiobutton(rule_frame, text="模糊匹配", variable=rule_var, value="fuzzy", font=("SimHei", 8)).pack(
            side=tk.LEFT)
        rule_frame.pack(side=tk.LEFT, padx=5)

        # 删除按钮
        def remove_pair():
            for i, pair in enumerate(self.match_pairs):
                if pair["frame"] == pair_frame:
                    self.match_pairs.pop(i)
                    break
            pair_frame.destroy()

        tk.Button(pair_frame, text="删除", command=remove_pair, font=self.small_font, bg="#f8f9fa").pack(side=tk.LEFT,
                                                                                                         padx=5)

        self.match_pairs.append({
            "col1_var": col1_var,
            "col1_combo": col1_combo,
            "col2_var": col2_var,
            "col2_combo": col2_combo,
            "rule_var": rule_var,
            "frame": pair_frame
        })

        self.update_match_combos()

    def update_match_combos(self):
        if self.df1 is not None:
            cols1 = list(self.df1.columns)
            for pair in self.match_pairs:
                pair["col1_combo"]["state"] = "readonly"
                pair["col1_combo"]["values"] = cols1
                if cols1 and not pair["col1_var"].get():
                    pair["col1_var"].set(cols1[0])

        if self.df2 is not None:
            cols2 = list(self.df2.columns)
            for pair in self.match_pairs:
                pair["col2_combo"]["state"] = "readonly"
                pair["col2_combo"]["values"] = cols2
                if cols2 and not pair["col2_var"].get():
                    pair["col2_var"].set(cols2[0])

    def browse_file1(self):
        path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx;*.xls")])
        if path:
            self.file1_path.set(path)
            self.load_file_info(path, 1)

    def browse_file2(self):
        path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx;*.xls")])
        if path:
            self.file2_path.set(path)
            self.load_file_info(path, 2)

    def browse_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel文件", "*.xlsx")])
        if path:
            self.output_path.set(path)

    def load_file_info(self, path, file_num):
        def load_info():
            try:
                df = pd.read_excel(path, nrows=100)
                info = f"文件名：{os.path.basename(path)}\n"
                info += f"路径：{path}\n"
                info += f"列数：{len(df.columns)}\n"
                total_rows = pd.read_excel(path, usecols=[0]).shape[0]
                info += f"行数：{total_rows}\n"
                info += "列名：\n"
                for i, col in enumerate(df.columns[:5]):
                    info += f"  第{i + 1}列：{col}\n"
                if len(df.columns) > 5:
                    info += f"  ... 共{len(df.columns)}列\n"

                if file_num == 1:
                    self.df1 = df
                    self.root.after(0, lambda: self._update_text(self.info1, info))
                else:
                    self.df2 = df
                    self.root.after(0, lambda: self._update_text(self.info2, info))
                    self.root.after(0, self.update_col_selection)

                self.root.after(0, self.update_match_combos)

            except Exception as e:
                self.root.after(0, lambda err=str(e): messagebox.showerror("错误", f"读取文件信息失败：{err}"))

        threading.Thread(target=load_info, daemon=True).start()

    def update_col_selection(self):
        self.col_listbox.delete(0, tk.END)
        if self.df2 is None:
            return

        for col in self.df2.columns:
            self.col_listbox.insert(tk.END, col)

        if len(self.df2.columns) >= 3:
            for i in range(3):
                self.col_listbox.selection_set(i)

    def select_all_cols(self):
        for i in range(self.col_listbox.size()):
            self.col_listbox.selection_set(i)

    def deselect_all_cols(self):
        self.col_listbox.selection_clear(0, tk.END)

    def _update_text(self, text_widget, content):
        text_widget.config(state=tk.NORMAL)
        text_widget.delete(1.0, tk.END)
        text_widget.insert(tk.END, content)
        text_widget.config(state=tk.DISABLED)

    def _append_log(self, content):
        self.log_text.config(state=tk.NORMAL)
        if int(self.log_text.index('end-1c').split('.')[0]) > 100:
            self.log_text.delete(1.0, 2.0)
        self.log_text.insert(tk.END, content + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def preview_file1(self):
        self._preview_file(self.file1_path.get(), "文件1预览")

    def preview_file2(self):
        self._preview_file(self.file2_path.get(), "文件2预览")

    def _preview_file(self, path, title):
        if not path or not os.path.exists(path):
            messagebox.showinfo("提示", "请先选择有效的文件")
            return

        def load_preview():
            try:
                df = pd.read_excel(path, nrows=50)
                self.root.after(0, lambda: self._create_preview_window(df, title))
            except Exception as e:
                self.root.after(0, lambda err=str(e): messagebox.showerror("错误", f"预览失败：{err}"))

        threading.Thread(target=load_preview, daemon=True).start()

    def _create_preview_window(self, df, title):
        preview_window = tk.Toplevel(self.root)
        preview_window.title(title)
        preview_window.geometry("800x500")

        tree = ttk.Treeview(preview_window, show="headings")
        columns = list(df.columns[:10])
        tree["columns"] = columns

        for col in columns:
            tree.heading(col, text=str(col))
            tree.column(col, width=100, anchor=tk.CENTER)

        for i, row in df.iterrows():
            values = [str(row[col]) if pd.notna(row[col]) else "" for col in columns]
            tree.insert("", tk.END, values=values)

        tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        vscrollbar = ttk.Scrollbar(tree, orient="vertical", command=tree.yview)
        tree.configure(yscroll=vscrollbar.set)
        vscrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        hscrollbar = ttk.Scrollbar(preview_window, orient="horizontal", command=tree.xview)
        tree.configure(xscroll=hscrollbar.set)
        hscrollbar.pack(fill=tk.X, padx=10)

    def clear_selection(self):
        self.file1_path.set("")
        self.file2_path.set("")
        self.output_path.set("合并结果.xlsx")
        self._update_text(self.info1, "")
        self._update_text(self.info2, "")
        self._update_text(self.log_text, "")
        self.col_listbox.delete(0, tk.END)

        for pair in self.match_pairs[1:]:
            pair["frame"].destroy()
        self.match_pairs = self.match_pairs[:1] if self.match_pairs else []
        if not self.match_pairs:
            self.add_match_pair()
        else:
            self.match_pairs[0]["col1_var"].set("")
            self.match_pairs[0]["col2_var"].set("")
            self.match_pairs[0]["rule_var"].set("fuzzy")
            self.match_pairs[0]["col1_combo"]["state"] = "disabled"
            self.match_pairs[0]["col2_combo"]["state"] = "disabled"

        self.status_var.set("就绪")
        self.time_var.set("耗时：--:--")
        self.progress["value"] = 0
        self.df1 = None
        self.df2 = None
        self.selected_cols = []

    def start_progress_listener(self):
        def listen():
            while True:
                try:
                    time.sleep(0.2)
                    if not self.is_running:
                        continue

                    if not self.progress_queue.empty():
                        msg = self.progress_queue.get()
                        if msg["type"] == "progress":
                            self.processed_rows = msg["processed"]
                            self.total_rows = msg["total"]

                            if self.total_rows > 0:
                                progress = 20 + int(70 * self.processed_rows / self.total_rows)
                                self.root.after(0, lambda p=progress: self.progress.configure(value=p))
                                self.root.after(0, lambda: self.status_var.set(
                                    f"处理第{self.processed_rows}/{self.total_rows}行..."
                                ))

                        elapsed = time.time() - self.start_time
                        minutes = int(elapsed // 60)
                        seconds = int(elapsed % 60)
                        self.root.after(0, lambda m=minutes, s=seconds: self.time_var.set(
                            f"耗时：{m:02d}:{s:02d}"
                        ))
                except Exception:
                    break

        threading.Thread(target=listen, daemon=True).start()

    def start_log_listener(self):
        def listen():
            while True:
                try:
                    time.sleep(0.1)
                    if not self.log_queue.empty():
                        log_msg = self.log_queue.get()
                        self.root.after(0, lambda msg=log_msg: self._append_log(msg))
                except Exception:
                    break

        threading.Thread(target=listen, daemon=True).start()

    def start_merge_process(self):
        if self.is_running:
            messagebox.showinfo("提示", "合并操作正在进行中，请稍后...")
            return

        file1 = self.file1_path.get()
        file2 = self.file2_path.get()
        output = self.output_path.get()

        if not (file1 and file2 and output):
            messagebox.showerror("错误", "请填写所有文件路径")
            return

        if not os.path.exists(file1) or not os.path.exists(file2):
            messagebox.showerror("错误", "文件不存在")
            return

        valid_pairs = []
        for i, pair in enumerate(self.match_pairs, 1):
            col1 = pair["col1_var"].get()
            col2 = pair["col2_var"].get()
            rule = pair["rule_var"].get()
            if not col1 or not col2:
                messagebox.showerror("错误", f"第{i}对匹配列未完整选择")
                return
            valid_pairs.append({"col1": col1, "col2": col2, "rule": rule})

        if not valid_pairs:
            messagebox.showerror("错误", "请至少添加一对匹配列")
            return

        self.selected_cols = [self.df2.columns[i] for i in self.col_listbox.curselection()]
        if not self.selected_cols:
            messagebox.showerror("错误", "请选择需要合并的列")
            return

        self._update_text(self.log_text, "")

        self.is_running = True
        self.start_time = time.time()
        self.processed_rows = 0
        self.run_btn.config(state=tk.DISABLED)
        self.cancel_btn.config(state=tk.NORMAL)
        self.status_var.set("正在初始化处理进程...")
        self.time_var.set("耗时：00:00")
        self.progress["value"] = 5

        self.merge_process = Process(
            target=process_merge,
            args=(
                file1, file2, output, valid_pairs, self.selected_cols,
                self.progress_queue, self.control_queue, self.log_queue
            )
        )
        self.merge_process.start()

        self.root.after(500, self.check_process_status)

    def check_process_status(self):
        if not self.is_running:
            return

        if self.merge_process.is_alive():
            if not self.control_queue.empty():
                cmd = self.control_queue.get()
                if cmd == "cancel":
                    self.merge_process.terminate()
                    self.is_running = False
                    self.status_var.set("合并已取消")
                    self.run_btn.config(state=tk.NORMAL)
                    self.cancel_btn.config(state=tk.DISABLED)
                    return
            self.root.after(500, self.check_process_status)
        else:
            self.is_running = False
            self.progress["value"] = 100
            self.status_var.set("合并完成！")
            self.run_btn.config(state=tk.NORMAL)
            self.cancel_btn.config(state=tk.DISABLED)

            elapsed = time.time() - self.start_time
            minutes = int(elapsed // 60)
            seconds = int(elapsed % 60)
            self.time_var.set(f"耗时：{minutes:02d}:{seconds:02d}")

            if messagebox.askyesno("成功", f"文件已保存至：\n{self.output_path.get()}\n是否打开？"):
                try:
                    os.startfile(self.output_path.get())
                except:
                    pass

    def cancel_merge(self):
        if not self.is_running:
            return

        if messagebox.askyesno("确认", "确定取消？"):
            self.control_queue.put("cancel")
            self.status_var.set("正在取消...")
            self.cancel_btn.config(state=tk.DISABLED)


def process_merge(file1, file2, output, match_pairs, selected_cols, progress_queue, control_queue, log_queue):
    """独立进程中的数据处理逻辑（强化多条件联合匹配）"""
    try:
        # 1. 读取数据
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)
        total_rows = len(df1)
        if total_rows == 0 or len(df2) == 0:
            raise ValueError("文件无数据")

        log_queue.put(f"开始多条件联合匹配（共{len(match_pairs)}个匹配条件）")

        # 2. 预处理匹配列并构建索引（每个匹配对单独构建索引）
        match_indexes = []
        for pair_idx, pair in enumerate(match_pairs, 1):
            col1, col2, rule = pair["col1"], pair["col2"], pair["rule"]

            # 处理表1列：清洗数据
            df1_col = df1[col1].fillna("").astype(str).str.strip()
            if rule == "fuzzy":
                df1_col = df1_col.str.lower()

            # 处理表2列：构建索引（存储所有匹配行的索引，而非仅第一个）
            df2_col = df2[col2].fillna("").astype(str).str.strip()
            if rule == "fuzzy":
                df2_col = df2_col.str.lower()

            # 索引结构：{值: [行索引1, 行索引2, ...]}（存储所有匹配行）
            index_data = {}
            for idx, val in df2_col.items():
                if val not in index_data:
                    index_data[val] = []
                index_data[val].append(idx)

            log_queue.put(f"匹配对{pair_idx}（{col1}→{col2}，{rule}）：表2共{len(index_data)}个唯一值")

            match_indexes.append({
                "pair_idx": pair_idx,
                "col1": col1,
                "col2": col2,
                "rule": rule,
                "df1_col": df1_col,
                "index_data": index_data,
                "df2": df2
            })

        # 3. 准备合并列
        target_cols = {col: f"{col}_from_file2" for col in selected_cols}
        for col in target_cols.values():
            df1[col] = None

        # 4. 处理数据（核心：多条件联合匹配）
        batch_size = 100
        for i in range(0, total_rows, batch_size):
            # 检查取消命令
            if not control_queue.empty() and control_queue.get() == "cancel":
                return

            end = min(i + batch_size, total_rows)
            for idx in range(i, end):
                row_log = [f"行{idx + 1}：开始多条件匹配"]
                all_matched = True
                matched_sets = []  # 存储每个匹配对的匹配行索引集合

                # 逐个匹配条件校验
                for data in match_indexes:
                    val1 = data["df1_col"].iloc[idx]
                    row_log.append(f"  匹配对{data['pair_idx']}：'{val1}'（规则：{data['rule']}）")

                    # 空值处理
                    if not val1:
                        all_matched = False
                        row_log.append(f"  → 空值，不满足该条件")
                        break

                    # 完全相同规则
                    if data["rule"] == "exact":
                        if val1 in data["index_data"]:
                            matched_indices = data["index_data"][val1]
                            matched_sets.append(set(matched_indices))  # 转为集合便于交集计算
                            row_log.append(f"  → 满足，表2匹配行：{matched_indices[:3]}...（共{len(matched_indices)}行）")
                        else:
                            all_matched = False
                            row_log.append(f"  → 不满足（无完全匹配值）")
                            break

                    # 模糊匹配规则
                    else:
                        matched_indices = []
                        for val2, idx2_list in data["index_data"].items():
                            if val1 in val2 or val2 in val1:
                                matched_indices.extend(idx2_list)
                        if matched_indices:
                            matched_sets.append(set(matched_indices))
                            row_log.append(f"  → 满足，表2匹配行：{matched_indices[:3]}...（共{len(matched_indices)}行）")
                        else:
                            all_matched = False
                            row_log.append(f"  → 不满足（无模糊匹配值）")
                            break

                # 所有条件都满足后，计算交集（必须存在共同匹配的行）
                if all_matched and matched_sets:
                    # 计算所有匹配对的行索引交集（找到同时满足所有条件的行）
                    common_indices = matched_sets[0]
                    for s in matched_sets[1:]:
                        common_indices.intersection_update(s)  # 交集运算

                    if common_indices:
                        # 取第一个共同匹配行
                        selected_idx = sorted(common_indices)[0]
                        matched_row = df2.iloc[selected_idx]

                        # 合并数据
                        for col in selected_cols:
                            df1.at[idx, target_cols[col]] = matched_row[col]

                        row_log.append(f"  → 所有条件均满足！合并表2行{selected_idx + 1}")
                        if len(common_indices) > 1:
                            row_log.append(f"  → 提示：存在{len(common_indices)}行同时满足所有条件，取第一行")
                    else:
                        all_matched = False
                        row_log.append(f"  → 各条件匹配行无交集，不合并")

                # 输出最终结果
                if all_matched:
                    row_log.append("  → 合并成功")
                else:
                    row_log.append("  → 未通过所有条件，不合并")

                # 输出日志（每10行）
                if idx % 10 == 0:
                    log_queue.put("\n".join(row_log))  # 换行显示，更清晰

            # 更新进度
            progress_queue.put({
                "type": "progress",
                "processed": end,
                "total": total_rows
            })

        # 5. 保存结果
        df1 = df1.dropna(axis=1, how='all')
        df1.to_excel(output, index=False)
        log_queue.put(f"合并完成，共处理{total_rows}行（仅保留所有条件都满足的记录）")

    except Exception as e:
        log_queue.put(f"错误：{str(e)}")
        progress_queue.put({"type": "error", "msg": str(e)})


