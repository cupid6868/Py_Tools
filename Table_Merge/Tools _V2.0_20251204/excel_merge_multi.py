import tkinter as tk
from tkinter import filedialog, messagebox, ttk, font
import pandas as pd
import os
import threading
from queue import Queue
import time
from typing import List, Dict
import numpy as np
import re
from datetime import datetime
import warnings

# 忽略Pandas未来的警告
warnings.filterwarnings("ignore", category=FutureWarning)

# 设置pandas选项，避免FutureWarning
pd.set_option('future.no_silent_downcasting', True)


class ExcelMergerPro:
    """主窗口 - 负责整体协调和结果显示"""

    def __init__(self, root):
        self.root = root
        self.root.title("数据智能合并工具 - 主控中心")

        # 设置窗口大小
        self.root.geometry("1200x800")
        self.root.minsize(1000, 700)
        self.root.state("zoomed")
        self.root.resizable(True, True)

        # 主题配色方案
        self.colors = {
            "primary": "#2c6ecb",
            "primary_light": "#4a89dc",
            "primary_dark": "#1e56a0",
            "secondary": "#f5f7fa",
            "accent": "#ff7043",
            "success": "#38b000",
            "warning": "#ffab00",
            "danger": "#e5383b",
            "text": "#2d3748",
            "text_light": "#718096",
            "border": "#e2e8f0",
            "hover": "#edf2f7",
            "card_bg": "#ffffff",
            "shadow": "#dee2e6",
            "selected": "#e8f4f8",
            "unselected": "#f8f9fa"
        }

        # 字体设置
        self.font = ("SimHei", 10)
        self.small_font = ("SimHei", 9)
        self.title_font = ("SimHei", 14, "bold")

        # 数据存储
        self.main_table_data = None  # 主表格数据
        self.sub_tables_data = []  # 子表格数据列表
        self.output_path = tk.StringVar(value="合并结果")
        self.output_format = tk.StringVar(value="xlsx")
        self.is_running = False
        self.start_time = 0
        self.merge_result_df = None

        # 线程通信
        self.progress_queue = Queue()
        self.control_queue = Queue()
        self.result_queue = Queue()

        # 进度更新相关
        self.last_progress_update = 0
        self.progress_update_interval = 0.1  # 100ms更新一次进度

        # 创建界面
        self.setup_styles()
        self.create_widgets()

        # 启动监听器
        self.start_progress_listener()
        self.start_result_listener()

    def setup_styles(self):
        """配置ttk控件样式"""
        style = ttk.Style()
        style.configure(".", font=self.font, background=self.colors["secondary"])
        style.configure("TFrame", background=self.colors["secondary"])
        style.configure("TButton",
                        background=self.colors["primary"],
                        foreground="white",
                        padding=6,
                        borderwidth=0)
        style.map("TButton",
                  background=[("active", self.colors["primary_light"]),
                              ("pressed", self.colors["primary_dark"])],
                  relief=[("pressed", tk.SUNKEN), ("active", tk.RAISED)])
        style.configure("TLabel",
                        background=self.colors["secondary"],
                        foreground=self.colors["text"])
        style.configure("TProgressbar",
                        troughcolor=self.colors["secondary"],
                        background=self.colors["primary"],
                        thickness=10)

    def create_widgets(self):
        """创建主窗口界面"""
        # 设置背景色
        self.root.configure(bg=self.colors["secondary"])

        # 主容器
        main_container = tk.Frame(self.root, bg=self.colors["secondary"])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # 标题区域
        title_frame = tk.Frame(main_container, bg=self.colors["primary"], height=70)
        title_frame.pack(fill=tk.X, pady=(0, 20))
        title_frame.pack_propagate(False)

        title_label = tk.Label(title_frame,
                               text="数据智能合并工具 - 主控中心",
                               font=("SimHei", 16, "bold"),
                               bg=self.colors["primary"],
                               fg="white")
        title_label.pack(pady=20)

        # 主内容区域（两列布局）
        content_frame = tk.Frame(main_container, bg=self.colors["secondary"])
        content_frame.pack(fill=tk.BOTH, expand=True)

        # 左侧控制面板 - 增大宽度为450
        left_panel = tk.Frame(content_frame, bg=self.colors["secondary"], width=450)
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 20))
        left_panel.pack_propagate(False)

        # 右侧结果展示区
        right_panel = tk.Frame(content_frame, bg=self.colors["secondary"])
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # ===== 左侧控制面板 =====
        control_card = tk.Frame(left_panel, bg=self.colors["card_bg"], bd=1, relief=tk.SOLID,
                                highlightbackground=self.colors["border"], padx=20, pady=15)
        control_card.pack(fill=tk.BOTH, expand=True)

        # 主表格设置
        tk.Label(control_card, text="📊 主表格设置", font=("SimHei", 12, "bold"),
                 bg=self.colors["card_bg"], fg=self.colors["primary"]).pack(anchor=tk.W, pady=(0, 10))

        # 主表格信息显示
        self.main_info_frame = tk.Frame(control_card, bg=self.colors["card_bg"])
        self.main_info_frame.pack(fill=tk.X, pady=(0, 15))

        self.main_info_label = tk.Label(self.main_info_frame, text="未设置主表格", font=self.font,
                                        bg=self.colors["card_bg"], fg=self.colors["text_light"])
        self.main_info_label.pack(anchor=tk.W)

        # 主表格操作按钮
        main_btn_frame = tk.Frame(control_card, bg=self.colors["card_bg"])
        main_btn_frame.pack(fill=tk.X, pady=(0, 20))

        tk.Button(main_btn_frame, text="设置主表格", command=self.open_main_table_window,
                  font=self.font, bg=self.colors["primary"], fg="white",
                  relief=tk.FLAT, padx=25, pady=6).pack(side=tk.LEFT, padx=(0, 10))
        self.add_hover_effect(main_btn_frame.winfo_children()[0], self.colors["primary"], self.colors["primary_light"])

        tk.Button(main_btn_frame, text="预览主表格", command=self.preview_main_table,
                  font=self.font, bg=self.colors["secondary"], fg=self.colors["text"],
                  relief=tk.FLAT, padx=25, pady=6).pack(side=tk.LEFT)
        self.add_hover_effect(main_btn_frame.winfo_children()[1], self.colors["secondary"], self.colors["hover"])

        # 子表格管理
        tk.Label(control_card, text="🗂️ 子表格管理", font=("SimHei", 12, "bold"),
                 bg=self.colors["card_bg"], fg=self.colors["accent"]).pack(anchor=tk.W, pady=(0, 10))

        # 子表格列表容器
        sub_list_frame = tk.Frame(control_card, bg=self.colors["secondary"], height=200)
        sub_list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        # 创建带滚动条的列表框
        list_container = tk.Frame(sub_list_frame, bg=self.colors["secondary"])
        list_container.pack(fill=tk.BOTH, expand=True)

        # 列表框
        self.sub_listbox = tk.Listbox(list_container, font=self.font, bg=self.colors["card_bg"],
                                      bd=1, relief=tk.SOLID, selectbackground=self.colors["primary_light"])
        self.sub_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 滚动条
        list_scrollbar = tk.Scrollbar(list_container, orient=tk.VERTICAL, command=self.sub_listbox.yview)
        list_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.sub_listbox.config(yscrollcommand=list_scrollbar.set)

        # 子表格操作按钮
        sub_btn_frame = tk.Frame(control_card, bg=self.colors["card_bg"])
        sub_btn_frame.pack(fill=tk.X, pady=(0, 20))

        tk.Button(sub_btn_frame, text="添加子表格", command=self.open_sub_table_window,
                  font=self.font, bg=self.colors["accent"], fg="white",
                  relief=tk.FLAT, padx=20, pady=6).pack(side=tk.LEFT, padx=(0, 10))
        self.add_hover_effect(sub_btn_frame.winfo_children()[0], self.colors["accent"], "#e64a19")

        tk.Button(sub_btn_frame, text="编辑选中", command=self.edit_selected_sub_table,
                  font=self.font, bg=self.colors["primary_light"], fg="white",
                  relief=tk.FLAT, padx=20, pady=6).pack(side=tk.LEFT, padx=(0, 10))
        self.add_hover_effect(sub_btn_frame.winfo_children()[1], self.colors["primary_light"], self.colors["primary"])

        tk.Button(sub_btn_frame, text="删除选中", command=self.delete_selected_sub_table,
                  font=self.font, bg=self.colors["danger"], fg="white",
                  relief=tk.FLAT, padx=20, pady=6).pack(side=tk.LEFT)
        self.add_hover_effect(sub_btn_frame.winfo_children()[2], self.colors["danger"], "#c92a2a")

        # 输出设置
        tk.Label(control_card, text="💾 输出设置", font=("SimHei", 12, "bold"),
                 bg=self.colors["card_bg"], fg=self.colors["primary"]).pack(anchor=tk.W, pady=(0, 10))

        output_frame = tk.Frame(control_card, bg=self.colors["card_bg"])
        output_frame.pack(fill=tk.X, pady=(0, 20))

        tk.Label(output_frame, text="输出文件：", font=self.font, bg=self.colors["card_bg"]).pack(anchor=tk.W)

        output_entry_frame = tk.Frame(output_frame, bg=self.colors["card_bg"])
        output_entry_frame.pack(fill=tk.X, pady=(0, 10))

        output_entry = tk.Entry(output_entry_frame, textvariable=self.output_path, font=self.font,
                                bg=self.colors["secondary"], bd=1, relief=tk.SOLID)
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        tk.Button(output_entry_frame, text="浏览", command=self.browse_output,
                  font=self.small_font, bg=self.colors["primary"], fg="white",
                  relief=tk.FLAT, padx=10, pady=4).pack(side=tk.RIGHT)
        self.add_hover_effect(output_entry_frame.winfo_children()[1], self.colors["primary"],
                              self.colors["primary_light"])

        # 格式选择
        format_frame = tk.Frame(output_frame, bg=self.colors["card_bg"])
        format_frame.pack(fill=tk.X)

        tk.Label(format_frame, text="输出格式：", font=self.font, bg=self.colors["card_bg"]).pack(side=tk.LEFT)

        format_combo = ttk.Combobox(format_frame, textvariable=self.output_format,
                                    values=["xlsx", "csv", "dta"], state="readonly",
                                    width=12, font=self.small_font)
        format_combo.pack(side=tk.LEFT, padx=(10, 0))
        format_combo.current(0)

        # 合并控制
        tk.Label(control_card, text="⚙️ 合并控制", font=("SimHei", 12, "bold"),
                 bg=self.colors["card_bg"], fg=self.colors["success"]).pack(anchor=tk.W, pady=(0, 10))

        merge_btn_frame = tk.Frame(control_card, bg=self.colors["card_bg"])
        merge_btn_frame.pack(fill=tk.X)

        self.run_btn = tk.Button(merge_btn_frame, text="开始合并", command=self.start_merge,
                                 font=("SimHei", 12, "bold"), bg=self.colors["success"], fg="white",
                                 relief=tk.FLAT, padx=25, pady=8)
        self.run_btn.pack(side=tk.LEFT, padx=(0, 10))
        self.add_hover_effect(self.run_btn, self.colors["success"], "#4cc964")

        self.cancel_btn = tk.Button(merge_btn_frame, text="取消合并", command=self.cancel_merge,
                                    font=("SimHei", 12), bg=self.colors["warning"], fg="white",
                                    state=tk.DISABLED, relief=tk.FLAT, padx=25, pady=8)
        self.cancel_btn.pack(side=tk.LEFT)
        self.add_hover_effect(self.cancel_btn, self.colors["warning"], "#ffc145")

        # 状态信息
        status_frame = tk.Frame(control_card, bg=self.colors["card_bg"])
        status_frame.pack(fill=tk.X, pady=(20, 0))

        self.status_var = tk.StringVar(value="就绪")
        status_label = tk.Label(status_frame, textvariable=self.status_var, font=self.font,
                                bg=self.colors["card_bg"], fg=self.colors["text"])
        status_label.pack(anchor=tk.W)

        self.time_var = tk.StringVar(value="耗时：--:--")
        time_label = tk.Label(status_frame, textvariable=self.time_var, font=self.small_font,
                              bg=self.colors["card_bg"], fg=self.colors["text_light"])
        time_label.pack(anchor=tk.W)

        self.progress = ttk.Progressbar(status_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill=tk.X, pady=(5, 0))

        # ===== 右侧结果展示区 =====
        result_card = tk.Frame(right_panel, bg=self.colors["card_bg"], bd=1, relief=tk.SOLID,
                               highlightbackground=self.colors["border"])
        result_card.pack(fill=tk.BOTH, expand=True)

        # 结果标题
        tk.Label(result_card, text="📈 合并结果预览", font=("SimHei", 14, "bold"),
                 bg=self.colors["card_bg"], fg=self.colors["primary"]).pack(anchor=tk.W, padx=20, pady=15)

        # 创建结果表格容器
        tree_container = tk.Frame(result_card, bg=self.colors["secondary"], padx=20)
        tree_container.pack(fill=tk.BOTH, expand=True, pady=(0, 20))

        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        # 创建Treeview
        self.result_tree = ttk.Treeview(tree_container, show="headings")
        self.result_tree.grid(row=0, column=0, sticky=tk.NSEW)

        # 垂直滚动条
        vscrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=self.result_tree.yview)
        vscrollbar.grid(row=0, column=1, sticky=tk.NS)
        self.result_tree.configure(yscrollcommand=vscrollbar.set)

        # 水平滚动条
        hscrollbar = ttk.Scrollbar(result_card, orient="horizontal", command=self.result_tree.xview)
        hscrollbar.pack(fill=tk.X, padx=20, pady=(0, 15))
        self.result_tree.configure(xscrollcommand=hscrollbar.set)

        # 绑定滚轮事件
        self.result_tree.bind("<MouseWheel>", self._treeview_horizontal_wheel)

    def add_hover_effect(self, button, normal_bg, hover_bg):
        """添加按钮悬停效果"""

        def on_enter(e):
            button['background'] = hover_bg

        def on_leave(e):
            button['background'] = normal_bg

        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)

    def _treeview_horizontal_wheel(self, event):
        """处理Treeview横向滚轮"""
        if event.delta:
            self.result_tree.xview_scroll(int(-1 * (event.delta / 120)), "units")

    def open_main_table_window(self):
        """打开主表格设置窗口"""
        if self.is_running:
            messagebox.showinfo("提示", "请等待当前操作完成")
            return
        MainTableWindow(self.root, self)

    def open_sub_table_window(self, index=None):
        """打开子表格设置窗口"""
        if self.is_running:
            messagebox.showinfo("提示", "请等待当前操作完成")
            return

        if index is None:
            # 添加新的子表格
            SubTableWindow(self.root, self, None)
        else:
            # 编辑现有子表格
            if 0 <= index < len(self.sub_tables_data):
                SubTableWindow(self.root, self, self.sub_tables_data[index])

    def edit_selected_sub_table(self):
        """编辑选中的子表格"""
        selection = self.sub_listbox.curselection()
        if not selection:
            messagebox.showinfo("提示", "请先选择一个子表格")
            return
        index = selection[0]
        self.open_sub_table_window(index)

    def delete_selected_sub_table(self):
        """删除选中的子表格"""
        selection = self.sub_listbox.curselection()
        if not selection:
            messagebox.showinfo("提示", "请先选择一个子表格")
            return

        index = selection[0]
        if messagebox.askyesno("确认删除", f"确定删除子表格 {index + 1} 吗？"):
            self.sub_tables_data.pop(index)
            self.update_sub_table_list()

    def update_sub_table_list(self):
        """更新子表格列表显示"""
        self.sub_listbox.delete(0, tk.END)
        for i, sub_data in enumerate(self.sub_tables_data):
            filename = os.path.basename(sub_data["path"]) if sub_data["path"] else "未设置文件"
            match_count = len(sub_data["match_pairs"])
            selected_count = len(sub_data["selected_cols"])
            self.sub_listbox.insert(tk.END, f"子表格{i + 1}: {filename} (匹配{match_count}对, 选择{selected_count}列)")

    def update_main_table_info(self, data):
        """更新主表格信息"""
        self.main_table_data = data
        if data:
            filename = os.path.basename(data["path"])
            row_count = data["row_count"]
            col_count = data["col_count"]
            self.main_info_label.config(
                text=f"{filename}\n{row_count}行 × {col_count}列",
                fg=self.colors["text"]
            )
        else:
            self.main_info_label.config(text="未设置主表格", fg=self.colors["text_light"])

    def add_sub_table_data(self, data):
        """添加子表格数据"""
        self.sub_tables_data.append(data)
        self.update_sub_table_list()

    def update_sub_table_data(self, old_data, new_data):
        """更新子表格数据"""
        for i, sub_data in enumerate(self.sub_tables_data):
            if sub_data == old_data:
                self.sub_tables_data[i] = new_data
                break
        self.update_sub_table_list()

    def preview_main_table(self):
        """预览主表格"""
        if not self.main_table_data:
            messagebox.showinfo("提示", "请先设置主表格")
            return
        self._preview_file(self.main_table_data["path"], "主表格预览")

    def browse_output(self):
        """选择输出文件路径"""
        selected_format = self.output_format.get()
        filetypes = []
        default_ext = f".{selected_format}"

        if selected_format == "xlsx":
            filetypes = [("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        elif selected_format == "csv":
            filetypes = [("CSV文件", "*.csv"), ("所有文件", "*.*")]
        elif selected_format == "dta":
            filetypes = [("Stata文件", "*.dta"), ("所有文件", "*.*")]

        path = filedialog.asksaveasfilename(
            defaultextension=default_ext,
            filetypes=filetypes,
            title=f"保存{selected_format.upper()}文件"
        )
        if path:
            self.output_path.set(path)

    def start_merge(self):
        """开始合并操作"""
        if self.is_running:
            return

        # 验证数据
        if not self.main_table_data:
            messagebox.showerror("错误", "请先设置主表格")
            return

        if not self.sub_tables_data:
            messagebox.showerror("错误", "请至少添加一个子表格")
            return

        # 验证输出路径
        output_path = self.output_path.get()
        if not output_path:
            messagebox.showerror("错误", "请设置输出文件路径")
            return

        output_format = self.output_format.get()
        if not output_path.endswith(f".{output_format}"):
            output_path += f".{output_format}"
            self.output_path.set(output_path)

        # 准备合并数据
        merge_config = {
            "main_table": self.main_table_data,
            "sub_tables": self.sub_tables_data,
            "output_path": output_path,
            "output_format": output_format
        }

        # 启动合并线程
        self.is_running = True
        self.start_time = time.time()
        self.run_btn.config(state=tk.DISABLED)
        self.cancel_btn.config(state=tk.NORMAL)
        self.status_var.set("开始合并...")
        self.progress["value"] = 0
        self.last_progress_update = time.time()

        # 清空结果预览
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        merge_thread = threading.Thread(
            target=self._run_merge_process,
            args=(merge_config,),
            daemon=True
        )
        merge_thread.start()

    def _run_merge_process(self, config):
        """运行合并进程"""
        try:
            # 读取主表格完整数据
            main_df = self._read_full_data(config["main_table"]["path"])
            total_main_rows = len(main_df)
            if total_main_rows == 0:
                raise ValueError("主表格无有效数据")

            total_progress_units = total_main_rows * len(config["sub_tables"])
            current_progress = 0

            # 处理每个子表格
            for sub_idx, sub_table in enumerate(config["sub_tables"], 1):
                # 检查取消信号
                if not self.control_queue.empty() and self.control_queue.get() == "cancel":
                    self.result_queue.put(
                        {"type": "error", "msg": f"合并已取消 - 已完成子表格1~{sub_idx - 1}的合并", "df": main_df})
                    return

                # 发送状态
                current_time = time.time()
                if current_time - self.last_progress_update > self.progress_update_interval:
                    self.progress_queue.put({
                        "type": "progress",
                        "processed": current_progress,
                        "total": total_progress_units,
                        "status": f"正在处理子表格 {sub_idx}/{len(config['sub_tables'])}"
                    })
                    self.last_progress_update = current_time

                # 读取子表格完整数据
                sub_df = self._read_full_data(sub_table["path"])
                if len(sub_df) == 0:
                    # 即使跳过子表格也要更新进度
                    current_progress += total_main_rows
                    if current_progress > total_progress_units:
                        current_progress = total_progress_units
                    self.result_queue.put({"type": "warning", "msg": f"子表格 {sub_idx} 无有效数据，跳过"})
                    continue

                # 准备匹配索引
                match_indexes = []
                valid_match_flag = False

                for pair_idx, pair in enumerate(sub_table["match_pairs"], 1):
                    col1, col2, rule = pair["col1"], pair["col2"], pair["rule"]

                    # 列存在性验证
                    if col1 not in main_df.columns:
                        raise ValueError(f"子表格 {sub_idx} - 匹配对{pair_idx}：主表格无列 '{col1}'")
                    if col2 not in sub_df.columns:
                        raise ValueError(f"子表格 {sub_idx} - 匹配对{pair_idx}：子表格无列 '{col2}'")

                    # 匹配列数据预处理
                    main_col = main_df[col1].fillna("").astype(str).str.strip()
                    sub_col = sub_df[col2].fillna("").astype(str).str.strip()

                    if rule == "fuzzy":
                        main_col = main_col.str.lower()
                        sub_col = sub_col.str.lower()

                    # 构建子表匹配映射
                    index_data = {}
                    for idx, val in sub_col.items():
                        if val.strip() != "":
                            index_data.setdefault(val, []).append(idx)

                    if len(index_data) > 0:
                        valid_match_flag = True

                    match_indexes.append({
                        "col1": col1,
                        "col2": col2,
                        "rule": rule,
                        "main_col": main_col,
                        "index_data": index_data,
                        "sub_df": sub_df
                    })

                # 子表无可用匹配数据时，跳过合并
                if not valid_match_flag:
                    current_progress += total_main_rows
                    if current_progress > total_progress_units:
                        current_progress = total_progress_units
                    self.result_queue.put({"type": "warning", "msg": f"子表格 {sub_idx} 无有效匹配列数据，跳过"})
                    continue

                # 准备合并列（避免重名）
                target_cols = {}
                for col in sub_table["selected_cols"]:
                    target_col_name = f"{col}_from_sub{sub_idx}"
                    suffix = 1
                    while target_col_name in main_df.columns:
                        target_col_name = f"{col}_from_sub{sub_idx}_{suffix}"
                        suffix += 1
                    target_cols[col] = target_col_name
                    main_df[target_col_name] = None  # 初始化合并列为空

                # 执行合并
                sub_total_rows = len(main_df)
                matched_count = 0
                batch_size = 10  # 减小批次大小以获得更频繁的更新

                for i in range(0, sub_total_rows, batch_size):
                    end = min(i + batch_size, sub_total_rows)
                    for idx in range(i, end):
                        all_matched = True
                        matched_sets = []

                        # 多条件匹配
                        for data in match_indexes:
                            val1 = data["main_col"].iloc[idx]
                            if val1.strip() == "":
                                all_matched = False
                                break

                            if data["rule"] == "exact":
                                if val1 not in data["index_data"]:
                                    all_matched = False
                                    break
                                matched_sets.append(set(data["index_data"][val1]))
                            else:  # fuzzy
                                matched_ids = []
                                for val2, ids in data["index_data"].items():
                                    if val1 in val2 or val2 in val1:
                                        matched_ids.extend(ids)
                                if not matched_ids:
                                    all_matched = False
                                    break
                                matched_sets.append(set(matched_ids))

                        # 匹配成功则写入数据
                        if all_matched and matched_sets:
                            common_ids = set.intersection(*matched_sets)
                            if common_ids:
                                match_id = next(iter(common_ids))
                                for col, target_col in target_cols.items():
                                    main_df.at[idx, target_col] = sub_df.at[match_id, col]
                                matched_count += 1

                    # 更新进度
                    current_progress += (end - i)

                    # 确保进度值不超过总进度
                    if current_progress > total_progress_units:
                        current_progress = total_progress_units

                    # 更频繁地发送进度更新
                    current_time = time.time()
                    if current_time - self.last_progress_update > self.progress_update_interval:
                        self.progress_queue.put({
                            "type": "progress",
                            "processed": current_progress,
                            "total": total_progress_units,
                            "status": f"子表格 {sub_idx} 处理中（已匹配{matched_count}行）"
                        })
                        self.last_progress_update = current_time

                self.result_queue.put({"type": "status", "msg": f"子表格 {sub_idx} 合并完成（匹配{matched_count}行）"})

            # 合并完成后强制发送100%进度
            self.progress_queue.put({
                "type": "progress",
                "processed": total_progress_units,
                "total": total_progress_units,
                "status": "所有子表格合并完成"
            })

            # 强制传递结果数据
            self.result_queue.put({"type": "result", "df": main_df})

            # 保存结果
            self.progress_queue.put({
                "type": "progress",
                "processed": total_progress_units,
                "total": total_progress_units,
                "status": "正在保存结果文件..."
            })

            save_success = self._save_output_file(main_df, config["output_path"], config["output_format"])

            if save_success:
                self.result_queue.put({"type": "save_success", "path": config["output_path"], "df": main_df})
            else:
                self.result_queue.put({"type": "save_error", "msg": "文件保存失败", "df": main_df})

        except Exception as e:
            error_msg = f"合并失败：{str(e)}"
            self.result_queue.put({"type": "error", "msg": error_msg})

    def cancel_merge(self):
        """取消合并操作"""
        if self.is_running:
            if messagebox.askyesno("确认", "确定取消合并操作吗？"):
                self.control_queue.put("cancel")
                self.status_var.set("正在取消...")
                self.cancel_btn.config(state=tk.DISABLED)

    def _read_full_data(self, path):
        """读取完整数据文件"""
        ext = os.path.splitext(path)[1].lower()
        try:
            if ext in ['.xlsx', '.xls']:
                df = pd.read_excel(path)
            elif ext == '.dta':
                df = pd.read_stata(path)
            elif ext == '.csv':
                encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1']
                for encoding in encodings:
                    try:
                        df = pd.read_csv(path, encoding=encoding)
                        break
                    except:
                        continue
                else:
                    raise ValueError("CSV文件编码无法识别")
            else:
                raise ValueError(f"不支持的文件格式：{ext}")

            df = df.dropna(how='all')
            return df
        except Exception as e:
            raise Exception(f"读取文件失败：{str(e)}")

    def _save_output_file(self, df, path, output_format):
        """保存输出文件"""
        try:
            # 先创建数据副本，避免修改原始合并结果
            df_save = df.copy()

            # 统一处理空值
            for col in df_save.columns:
                dtype = str(df_save[col].dtype).lower()
                if 'object' in dtype or 'string' in dtype:
                    # 修复FutureWarning问题
                    df_save[col] = df_save[col].fillna('').infer_objects(copy=False).astype(str).str[:500]
                elif 'int' in dtype or 'float' in dtype:
                    df_save[col] = df_save[col].replace([np.inf, -np.inf], np.nan).fillna(0)
                elif 'datetime' in dtype:
                    df_save[col] = df_save[col].fillna(pd.NaT)
                elif 'bool' in dtype:
                    df_save[col] = df_save[col].fillna(False).astype(int)

            # 保存文件
            if output_format == "dta":
                df_save = self._clean_data_for_dta(df_save)
                df_save.to_stata(path, version=114)
            elif output_format == "xlsx":
                df_save.to_excel(path, index=False, engine='openpyxl')
            elif output_format == "csv":
                df_save.to_csv(path, index=False, encoding='utf-8-sig', sep=',', na_rep='')
            else:
                raise ValueError(f"不支持的输出格式：{output_format}")

            return True
        except Exception as e:
            raise Exception(f"保存{output_format.upper()}文件失败：{str(e)}")

    def _clean_data_for_dta(self, df):
        """清理数据以适配Stata格式"""
        df_clean = df.copy()
        df_clean.columns = df_clean.columns.astype(str)
        df_clean.columns = df_clean.columns.str.replace(r'[^\w\s]', '_', regex=True)
        df_clean.columns = df_clean.columns.str.replace(r'\s+', '_', regex=True)
        df_clean.columns = df_clean.columns.str[:32]

        if len(df_clean.columns) != len(set(df_clean.columns)):
            df_clean.columns = [f"col_{i}" for i in range(len(df_clean.columns))]

        for col in df_clean.columns:
            if df_clean[col].dtype == object:
                # 修复FutureWarning问题
                df_clean[col] = df_clean[col].fillna('').infer_objects(copy=False).astype(str).str.encode('utf-8',
                                                                                                          errors='replace').str.decode(
                    'latin-1').str[:244]
            elif 'datetime64' in str(df_clean[col].dtype):
                df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce').dt.date
            elif str(df_clean[col].dtype).startswith('category'):
                df_clean[col] = df_clean[col].astype(str).fillna('').infer_objects(copy=False).str.encode('utf-8',
                                                                                                          errors='replace').str.decode(
                    'latin-1')
            elif df_clean[col].dtype == bool:
                df_clean[col] = df_clean[col].astype(int)
            elif 'int' in str(df_clean[col].dtype).lower() or 'float' in str(df_clean[col].dtype).lower():
                df_clean[col] = df_clean[col].replace([np.inf, -np.inf], np.nan).fillna(0)

        return df_clean

    def start_progress_listener(self):
        """启动进度监听器 - 改进版本"""

        def listen():
            while True:
                try:
                    # 不睡眠，持续检查队列
                    if not self.progress_queue.empty():
                        msg = self.progress_queue.get()
                        if msg["type"] == "progress":
                            # 确保进度值正确处理
                            processed = msg["processed"]
                            total = msg["total"]

                            if total > 0:
                                progress = int(100 * processed / total)
                                # 确保进度不超过100%
                                if progress > 100:
                                    progress = 100

                                # 直接在主线程更新进度
                                self.root.after(0, lambda p=progress, s=msg["status"]: self._update_progress_ui(p, s))
                    else:
                        # 队列为空时短暂休息
                        time.sleep(0.01)

                except Exception as e:
                    print(f"进度监听器错误：{str(e)}")
                    time.sleep(0.1)

        threading.Thread(target=listen, daemon=True).start()

    def _update_progress_ui(self, progress_value, status_text):
        """更新进度条UI"""
        self.progress.configure(value=progress_value)
        self.status_var.set(f"{status_text} - 进度: {progress_value}%")

        # 更新耗时显示
        if self.is_running:
            elapsed = time.time() - self.start_time
            minutes = int(elapsed // 60)
            seconds = int(elapsed % 60)
            self.time_var.set(f"耗时：{minutes:02d}:{seconds:02d}")

    def start_result_listener(self):
        """启动结果监听器"""

        def listen():
            while True:
                try:
                    if not self.result_queue.empty():
                        msg = self.result_queue.get()
                        self.root.after(0, lambda m=msg: self.handle_result(m))

                    time.sleep(0.1)
                except:
                    break

        threading.Thread(target=listen, daemon=True).start()

    def handle_result(self, msg):
        """处理结果消息"""

        if msg["type"] == "result":
            self.merge_result_df = msg["df"]
            self.display_result(self.merge_result_df)
            self.status_var.set("合并完成")

        elif msg["type"] == "save_success":
            self.merge_result_df = msg["df"]
            self.display_result(self.merge_result_df)
            self.status_var.set(f"结果已保存：{msg['path']}")
            messagebox.showinfo("成功", f"合并结果已保存到：\n{msg['path']}")

        elif msg["type"] == "save_error":
            self.status_var.set("保存失败")
            messagebox.showerror("错误", msg["msg"])

        elif msg["type"] == "error":
            self.status_var.set("合并失败")
            messagebox.showerror("错误", msg["msg"])

        elif msg["type"] == "warning":
            messagebox.showwarning("警告", msg["msg"])

        elif msg["type"] == "status":
            self.status_var.set(msg["msg"])

        # 重置按钮状态
        self.is_running = False
        self.run_btn.config(state=tk.NORMAL)
        self.cancel_btn.config(state=tk.DISABLED)

    def display_result(self, df):
        """显示合并结果"""

        # 清空现有内容
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        # 设置列
        self.result_tree["columns"] = list(df.columns)
        for col in df.columns:
            self.result_tree.heading(col, text=col)
            # 自动调整列宽
            col_width = max(80, len(str(col)) * 10)
            self.result_tree.column(col, width=col_width, anchor=tk.W, minwidth=50)

        # 添加数据（最多显示100行）
        display_df = df.head(100)
        for i, row in display_df.iterrows():
            values = []
            for col in df.columns:
                val = row[col]
                if pd.isna(val):
                    values.append("")
                elif isinstance(val, (int, float)):
                    values.append(f"{val:.2f}")
                elif isinstance(val, datetime):
                    values.append(val.strftime("%Y-%m-%d"))
                else:
                    values.append(str(val))
            self.result_tree.insert("", tk.END, values=values)

    def _preview_file(self, path, title="预览"):
        """预览文件"""

        def load_preview():
            try:
                df = self._read_full_data(path)
                # 限制行数
                preview_df = df.head(50)
                self.root.after(0, lambda: PreviewWindow(self.root, preview_df, title))
            except Exception as e:
                self.root.after(0, lambda err=str(e): messagebox.showerror("错误", f"预览失败：{err}"))

        threading.Thread(target=load_preview, daemon=True).start()


class MainTableWindow:
    """主表格设置窗口"""

    def __init__(self, parent, main_app, existing_data=None):
        self.parent = parent
        self.main_app = main_app
        self.existing_data = existing_data

        # 创建窗口
        self.window = tk.Toplevel(parent)
        self.window.title("设置主表格")
        self.window.geometry("800x600")
        self.window.resizable(True, True)
        self.window.transient(parent)
        self.window.grab_set()

        # 配置窗口关闭事件
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)

        # 配色
        self.colors = main_app.colors
        self.font = main_app.font
        self.small_font = main_app.small_font

        # 数据
        self.file_path = tk.StringVar()
        self.df = None

        # 创建界面
        self.create_widgets()

        # 如果编辑现有数据，加载数据
        if self.existing_data:
            self.file_path.set(self.existing_data["path"])
            self.load_file_info()

    def create_widgets(self):
        """创建窗口界面"""
        self.window.configure(bg=self.colors["secondary"])

        # 主容器
        main_container = tk.Frame(self.window, bg=self.colors["secondary"])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # 标题
        title_frame = tk.Frame(main_container, bg=self.colors["primary"], height=50)
        title_frame.pack(fill=tk.X, pady=(0, 20))
        title_frame.pack_propagate(False)

        tk.Label(title_frame, text="设置主表格（合并基准）",
                 font=("SimHei", 14, "bold"), bg=self.colors["primary"], fg="white").pack(pady=10)

        # 文件选择区域
        file_frame = tk.Frame(main_container, bg=self.colors["card_bg"], padx=15, pady=15, bd=1, relief=tk.SOLID)
        file_frame.pack(fill=tk.X, pady=(0, 20))

        tk.Label(file_frame, text="选择主表格文件：", font=self.font,
                 bg=self.colors["card_bg"]).pack(anchor=tk.W, pady=(0, 10))

        # 文件路径输入和浏览按钮
        path_frame = tk.Frame(file_frame, bg=self.colors["card_bg"])
        path_frame.pack(fill=tk.X, pady=(0, 15))

        entry = tk.Entry(path_frame, textvariable=self.file_path, font=self.font,
                         bg=self.colors["secondary"], bd=1, relief=tk.SOLID)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        browse_btn = tk.Button(path_frame, text="浏览", command=self.browse_file,
                               font=self.font, bg=self.colors["primary"], fg="white",
                               relief=tk.FLAT, padx=15, pady=5)
        browse_btn.pack(side=tk.RIGHT)
        self.add_hover_effect(browse_btn, self.colors["primary"], self.colors["primary_light"])

        # 信息显示区域
        info_frame = tk.Frame(main_container, bg=self.colors["card_bg"], padx=15, pady=15, bd=1, relief=tk.SOLID)
        info_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))

        tk.Label(info_frame, text="表格信息：", font=self.font,
                 bg=self.colors["card_bg"]).pack(anchor=tk.W, pady=(0, 10))

        self.info_text = tk.Text(info_frame, font=self.small_font, state=tk.DISABLED,
                                 bg=self.colors["secondary"], bd=1, relief=tk.SOLID,
                                 wrap=tk.WORD, height=10)
        self.info_text.pack(fill=tk.BOTH, expand=True)

        # 按钮区域
        btn_frame = tk.Frame(main_container, bg=self.colors["secondary"])
        btn_frame.pack(fill=tk.X)

        preview_btn = tk.Button(btn_frame, text="预览表格", command=self.preview_table,
                                font=self.font, bg=self.colors["secondary"], fg=self.colors["text"],
                                relief=tk.FLAT, padx=20, pady=8)
        preview_btn.pack(side=tk.LEFT, padx=(0, 10))
        self.add_hover_effect(preview_btn, self.colors["secondary"], self.colors["hover"])

        ok_btn = tk.Button(btn_frame, text="确定", command=self.save_and_close,
                           font=self.font, bg=self.colors["success"], fg="white",
                           relief=tk.FLAT, padx=20, pady=8)
        ok_btn.pack(side=tk.RIGHT)
        self.add_hover_effect(ok_btn, self.colors["success"], "#4cc964")

        cancel_btn = tk.Button(btn_frame, text="取消", command=self.on_close,
                               font=self.font, bg=self.colors["danger"], fg="white",
                               relief=tk.FLAT, padx=20, pady=8)
        cancel_btn.pack(side=tk.RIGHT, padx=(0, 10))
        self.add_hover_effect(cancel_btn, self.colors["danger"], "#c92a2a")

    def add_hover_effect(self, button, normal_bg, hover_bg):
        """添加按钮悬停效果"""

        def on_enter(e):
            button['background'] = hover_bg

        def on_leave(e):
            button['background'] = normal_bg

        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)

    def browse_file(self):
        """浏览文件"""
        filetypes = [
            ("数据文件", "*.xlsx;*.xls;*.dta;*.csv"),
            ("Excel文件", "*.xlsx;*.xls"),
            ("Stata文件", "*.dta"),
            ("CSV文件", "*.csv"),
            ("所有文件", "*.*")
        ]

        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            self.file_path.set(path)
            self.load_file_info()

    def load_file_info(self):
        """加载文件信息"""

        def load():
            try:
                path = self.file_path.get()
                if not os.path.exists(path):
                    self.window.after(0, lambda: messagebox.showerror("错误", "文件不存在"))
                    return

                # 读取数据
                df = self._read_data_file(path, nrows=100)
                full_df = self._read_data_file(path)

                # 更新界面
                self.window.after(0, lambda: self._update_info(path, df, full_df))

            except Exception as e:
                self.window.after(0, lambda err=str(e): messagebox.showerror("错误", f"读取文件失败：{err}"))

        threading.Thread(target=load, daemon=True).start()

    def _read_data_file(self, path, nrows=None):
        """读取数据文件"""
        ext = os.path.splitext(path)[1].lower()
        try:
            if ext in ['.xlsx', '.xls']:
                df = pd.read_excel(path, nrows=nrows) if nrows else pd.read_excel(path)
            elif ext == '.dta':
                df = pd.read_stata(path)
                if nrows:
                    df = df.head(nrows)
            elif ext == '.csv':
                encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1']
                for encoding in encodings:
                    try:
                        if nrows:
                            df = pd.read_csv(path, encoding=encoding, nrows=nrows)
                        else:
                            df = pd.read_csv(path, encoding=encoding)
                        break
                    except:
                        continue
                else:
                    raise ValueError("CSV文件编码无法识别")
            else:
                raise ValueError(f"不支持的文件格式：{ext}")

            df = df.dropna(how='all')
            return df
        except Exception as e:
            raise Exception(f"读取文件失败：{str(e)}")

    def _update_info(self, path, df, full_df):
        """更新信息显示"""
        info = f"文件路径：{path}\n"
        info += f"文件格式：{os.path.splitext(path)[1].upper()}\n"
        info += f"预览行数：{len(df)} 行\n"
        info += f"实际行数：{len(full_df)} 行\n"
        info += f"列数：{len(df.columns)} 列\n\n"
        info += "列名列表：\n"

        for i, col in enumerate(df.columns):
            info += f"  {i + 1}. {col}\n"
            if i >= 20:  # 最多显示20列
                info += f"  ... 共{len(df.columns)}列\n"
                break

        # 显示数据类型
        info += "\n数据类型：\n"
        for i, (col, dtype) in enumerate(df.dtypes.items()):
            info += f"  {col}: {dtype}\n"
            if i >= 10:  # 最多显示10个数据类型
                info += f"  ...\n"
                break

        self.info_text.config(state=tk.NORMAL)
        self.info_text.delete(1.0, tk.END)
        self.info_text.insert(tk.END, info)
        self.info_text.config(state=tk.DISABLED)

        # 保存数据
        self.df = df
        self.full_df = full_df

    def preview_table(self):
        """预览表格"""
        path = self.file_path.get()
        if not path or not os.path.exists(path):
            messagebox.showinfo("提示", "请先选择有效的文件")
            return

        def load_preview():
            try:
                df = self._read_data_file(path, nrows=50)
                self.window.after(0, lambda: PreviewWindow(self.window, df, "主表格预览"))
            except Exception as e:
                self.window.after(0, lambda err=str(e): messagebox.showerror("错误", f"预览失败：{err}"))

        threading.Thread(target=load_preview, daemon=True).start()

    def save_and_close(self):
        """保存设置并关闭窗口"""
        path = self.file_path.get()
        if not path or not os.path.exists(path):
            messagebox.showerror("错误", "请选择有效的文件")
            return

        if self.df is None:
            messagebox.showerror("错误", "请等待文件信息加载完成")
            return

        # 准备数据
        data = {
            "path": path,
            "df": self.df,
            "full_df": self.full_df,
            "row_count": len(self.full_df),
            "col_count": len(self.df.columns),
            "columns": list(self.df.columns)
        }

        # 通知主窗口
        self.main_app.update_main_table_info(data)

        # 关闭窗口
        self.window.destroy()

    def on_close(self):
        """关闭窗口"""
        if messagebox.askyesno("确认", "确定取消设置主表格吗？"):
            self.window.destroy()


class SubTableWindow:
    """子表格设置窗口"""

    def __init__(self, parent, main_app, existing_data=None):
        self.parent = parent
        self.main_app = main_app
        self.existing_data = existing_data

        # 创建窗口
        self.window = tk.Toplevel(parent)
        self.window.title("设置子表格" if not existing_data else "编辑子表格")
        self.window.geometry("900x700")
        self.window.resizable(True, True)
        self.window.transient(parent)
        self.window.grab_set()

        # 配置窗口关闭事件
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)

        # 配色
        self.colors = main_app.colors
        self.font = main_app.font
        self.small_font = main_app.small_font

        # 数据
        self.file_path = tk.StringVar()
        self.df = None
        self.full_df = None
        self.match_pairs = []  # 匹配对列表
        self.selected_cols = set()  # 选择的列

        # 创建界面
        self.create_widgets()

        # 如果编辑现有数据，加载数据
        if self.existing_data:
            self.load_existing_data()

    def create_widgets(self):
        """创建窗口界面"""
        self.window.configure(bg=self.colors["secondary"])

        # 主容器（带滚动条）
        main_container = tk.Frame(self.window, bg=self.colors["secondary"])
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 创建Canvas和滚动条
        canvas = tk.Canvas(main_container, bg=self.colors["secondary"], highlightthickness=0)
        scrollbar = tk.Scrollbar(main_container, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.colors["secondary"])

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # 只在canvas上绑定鼠标滚轮事件
        canvas.bind("<MouseWheel>", lambda e: self._on_canvas_wheel(e, canvas))

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 标题
        title_frame = tk.Frame(scrollable_frame, bg=self.colors["primary"], height=50)
        title_frame.pack(fill=tk.X, pady=(0, 15))
        title_frame.pack_propagate(False)

        title_text = "设置子表格" if not self.existing_data else f"编辑子表格 {self.existing_data.get('index', '')}"
        tk.Label(title_frame, text=title_text,
                 font=("SimHei", 14, "bold"), bg=self.colors["primary"], fg="white").pack(pady=10)

        # 文件选择区域
        file_frame = tk.Frame(scrollable_frame, bg=self.colors["card_bg"], padx=15, pady=15, bd=1, relief=tk.SOLID)
        file_frame.pack(fill=tk.X, pady=(0, 15))

        tk.Label(file_frame, text="选择子表格文件：", font=self.font,
                 bg=self.colors["card_bg"]).pack(anchor=tk.W, pady=(0, 10))

        # 文件路径输入和浏览按钮
        path_frame = tk.Frame(file_frame, bg=self.colors["card_bg"])
        path_frame.pack(fill=tk.X, pady=(0, 15))

        entry = tk.Entry(path_frame, textvariable=self.file_path, font=self.font,
                         bg=self.colors["secondary"], bd=1, relief=tk.SOLID)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        browse_btn = tk.Button(path_frame, text="浏览", command=self.browse_file,
                               font=self.font, bg=self.colors["primary"], fg="white",
                               relief=tk.FLAT, padx=15, pady=5)
        browse_btn.pack(side=tk.RIGHT)
        self.add_hover_effect(browse_btn, self.colors["primary"], self.colors["primary_light"])

        # 信息显示区域
        info_frame = tk.Frame(scrollable_frame, bg=self.colors["card_bg"], padx=15, pady=15, bd=1, relief=tk.SOLID)
        info_frame.pack(fill=tk.X, pady=(0, 15))

        tk.Label(info_frame, text="表格信息：", font=self.font,
                 bg=self.colors["card_bg"]).pack(anchor=tk.W, pady=(0, 10))

        self.info_text = tk.Text(info_frame, font=self.small_font, state=tk.DISABLED,
                                 bg=self.colors["secondary"], bd=1, relief=tk.SOLID,
                                 wrap=tk.WORD, height=6)
        self.info_text.pack(fill=tk.X)

        # 匹配条件设置区域
        match_frame = tk.Frame(scrollable_frame, bg=self.colors["card_bg"], padx=15, pady=15, bd=1, relief=tk.SOLID)
        match_frame.pack(fill=tk.X, pady=(0, 15))

        tk.Label(match_frame, text="匹配条件设置（所有条件必须同时满足）：",
                 font=self.font, bg=self.colors["card_bg"]).pack(anchor=tk.W, pady=(0, 10))

        # 匹配对列表容器
        self.match_pairs_frame = tk.Frame(match_frame, bg=self.colors["card_bg"])
        self.match_pairs_frame.pack(fill=tk.X, pady=(0, 10))

        # 添加匹配对按钮
        add_pair_btn = tk.Button(match_frame, text="添加匹配对", command=self.add_match_pair,
                                 font=self.font, bg=self.colors["primary_light"], fg="white",
                                 relief=tk.FLAT, padx=10, pady=5)
        add_pair_btn.pack(anchor=tk.W)
        self.add_hover_effect(add_pair_btn, self.colors["primary_light"], self.colors["primary"])

        # 列选择区域
        col_frame = tk.Frame(scrollable_frame, bg=self.colors["card_bg"], padx=15, pady=15, bd=1, relief=tk.SOLID)
        col_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        tk.Label(col_frame, text="选择要合并的列（可多选）：",
                 font=self.font, bg=self.colors["card_bg"]).pack(anchor=tk.W, pady=(0, 10))

        # 列按钮容器（带滚动条）
        col_container = tk.Frame(col_frame, bg=self.colors["secondary"])
        col_container.pack(fill=tk.BOTH, expand=True)

        # 创建Canvas用于列按钮
        col_canvas = tk.Canvas(col_container, bg=self.colors["secondary"], highlightthickness=0)
        col_scrollbar = tk.Scrollbar(col_container, orient=tk.VERTICAL, command=col_canvas.yview)
        self.col_buttons_frame = tk.Frame(col_canvas, bg=self.colors["secondary"])

        self.col_buttons_frame.bind(
            "<Configure>",
            lambda e: col_canvas.configure(scrollregion=col_canvas.bbox("all"))
        )

        col_canvas.create_window((0, 0), window=self.col_buttons_frame, anchor="nw")
        col_canvas.configure(yscrollcommand=col_scrollbar.set)

        # 只在列按钮canvas上绑定鼠标滚轮事件
        col_canvas.bind("<MouseWheel>", lambda e: self._on_canvas_wheel(e, col_canvas))

        col_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        col_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 列选择操作按钮
        col_btn_frame = tk.Frame(col_frame, bg=self.colors["card_bg"])
        col_btn_frame.pack(fill=tk.X, pady=(10, 0))

        tk.Button(col_btn_frame, text="全选", command=self.select_all_cols,
                  font=self.small_font, bg=self.colors["secondary"], fg=self.colors["text"],
                  relief=tk.FLAT, padx=10, pady=3).pack(side=tk.LEFT, padx=(0, 5))

        tk.Button(col_btn_frame, text="清空", command=self.clear_all_cols,
                  font=self.small_font, bg=self.colors["secondary"], fg=self.colors["text"],
                  relief=tk.FLAT, padx=10, pady=3).pack(side=tk.LEFT, padx=(0, 5))

        # 已选择列显示
        self.selected_cols_var = tk.StringVar(value="已选择：0列")
        tk.Label(col_btn_frame, textvariable=self.selected_cols_var,
                 font=self.small_font, bg=self.colors["card_bg"]).pack(side=tk.RIGHT)

        # 按钮区域
        btn_frame = tk.Frame(scrollable_frame, bg=self.colors["secondary"])
        btn_frame.pack(fill=tk.X, pady=(10, 0))

        preview_btn = tk.Button(btn_frame, text="预览表格", command=self.preview_table,
                                font=self.font, bg=self.colors["secondary"], fg=self.colors["text"],
                                relief=tk.FLAT, padx=20, pady=8)
        preview_btn.pack(side=tk.LEFT, padx=(0, 10))
        self.add_hover_effect(preview_btn, self.colors["secondary"], self.colors["hover"])

        ok_btn = tk.Button(btn_frame, text="确定", command=self.save_and_close,
                           font=self.font, bg=self.colors["success"], fg="white",
                           relief=tk.FLAT, padx=20, pady=8)
        ok_btn.pack(side=tk.RIGHT)
        self.add_hover_effect(ok_btn, self.colors["success"], "#4cc964")

        cancel_btn = tk.Button(btn_frame, text="取消", command=self.on_close,
                               font=self.font, bg=self.colors["danger"], fg="white",
                               relief=tk.FLAT, padx=20, pady=8)
        cancel_btn.pack(side=tk.RIGHT, padx=(0, 10))
        self.add_hover_effect(cancel_btn, self.colors["danger"], "#c92a2a")

        # 初始添加一个匹配对
        self.add_match_pair()

    def add_hover_effect(self, button, normal_bg, hover_bg):
        """添加按钮悬停效果"""

        def on_enter(e):
            button['background'] = hover_bg

        def on_leave(e):
            button['background'] = normal_bg

        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)

    def _on_canvas_wheel(self, event, canvas):
        """处理Canvas滚轮事件"""
        try:
            # 只在Canvas有效时执行滚动
            if canvas.winfo_exists():
                # 获取delta值（处理不同操作系统）
                delta = event.delta
                if event.num == 4:
                    delta = 120
                elif event.num == 5:
                    delta = -120

                # 滚动Canvas
                canvas.yview_scroll(int(-1 * (delta / 120)), "units")
        except:
            pass  # 忽略Canvas已销毁的情况

    def browse_file(self):
        """浏览文件"""
        path = filedialog.askopenfilename(
            filetypes=[
                ("数据文件", "*.xlsx;*.xls;*.dta;*.csv"),
                ("Excel文件", "*.xlsx;*.xls"),
                ("Stata文件", "*.dta"),
                ("CSV文件", "*.csv"),
                ("所有文件", "*.*")
            ]
        )
        if path:
            self.file_path.set(path)
            self.load_file_info()

    def load_file_info(self):
        """加载文件信息"""

        def load():
            try:
                path = self.file_path.get()
                if not os.path.exists(path):
                    self.window.after(0, lambda: messagebox.showerror("错误", "文件不存在"))
                    return

                # 读取数据
                df = self._read_data_file(path, nrows=100)
                full_df = self._read_data_file(path)

                # 更新界面
                self.window.after(0, lambda: self._update_info(path, df, full_df))

            except Exception as e:
                self.window.after(0, lambda err=str(e): messagebox.showerror("错误", f"读取文件失败：{err}"))

        threading.Thread(target=load, daemon=True).start()

    def _read_data_file(self, path, nrows=None):
        """读取数据文件"""
        ext = os.path.splitext(path)[1].lower()
        try:
            if ext in ['.xlsx', '.xls']:
                df = pd.read_excel(path, nrows=nrows) if nrows else pd.read_excel(path)
            elif ext == '.dta':
                df = pd.read_stata(path)
                if nrows:
                    df = df.head(nrows)
            elif ext == '.csv':
                encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1']
                for encoding in encodings:
                    try:
                        if nrows:
                            df = pd.read_csv(path, encoding=encoding, nrows=nrows)
                        else:
                            df = pd.read_csv(path, encoding=encoding)
                        break
                    except:
                        continue
                else:
                    raise ValueError("CSV文件编码无法识别")
            else:
                raise ValueError(f"不支持的文件格式：{ext}")

            df = df.dropna(how='all')
            return df
        except Exception as e:
            raise Exception(f"读取文件失败：{str(e)}")

    def _update_info(self, path, df, full_df):
        """更新信息显示"""
        info = f"文件路径：{path}\n"
        info += f"文件格式：{os.path.splitext(path)[1].upper()}\n"
        info += f"预览行数：{len(df)} 行\n"
        info += f"实际行数：{len(full_df)} 行\n"
        info += f"列数：{len(df.columns)} 列\n"

        self.info_text.config(state=tk.NORMAL)
        self.info_text.delete(1.0, tk.END)
        self.info_text.insert(tk.END, info)
        self.info_text.config(state=tk.DISABLED)

        # 保存数据
        self.df = df
        self.full_df = full_df

        # 更新列按钮
        self.update_col_buttons()

        # 更新匹配对中的列选择
        self.update_match_pairs_combos()

    def update_col_buttons(self):
        """更新列选择按钮"""
        # 清除现有按钮
        for widget in self.col_buttons_frame.winfo_children():
            widget.destroy()

        if self.df is None:
            return

        # 创建列按钮（每行4个）
        cols = list(self.df.columns)
        for i, col in enumerate(cols):
            row = i // 4
            col_num = i % 4

            # 创建按钮
            btn = tk.Button(self.col_buttons_frame, text=col,
                            font=self.small_font, bg=self.colors["unselected"], fg=self.colors["text"],
                            relief=tk.RAISED, padx=5, pady=3, width=15,
                            command=lambda c=col: self.toggle_column(c))
            btn.grid(row=row, column=col_num, padx=5, pady=3, sticky=tk.W)

            # 如果之前已选择，更新状态
            if col in self.selected_cols:
                btn.config(bg=self.colors["primary"], fg="white", relief=tk.SUNKEN)

            self.add_hover_effect(btn,
                                  self.colors["primary"] if col in self.selected_cols else self.colors["unselected"],
                                  self.colors["primary_light"])

        # 更新已选择列显示
        self.update_selected_cols_display()

    def toggle_column(self, col_name):
        """切换列选择状态"""
        if col_name in self.selected_cols:
            self.selected_cols.remove(col_name)
        else:
            self.selected_cols.add(col_name)

        # 更新按钮状态
        self.update_col_buttons()

    def select_all_cols(self):
        """选择所有列"""
        if self.df is not None:
            self.selected_cols = set(self.df.columns)
            self.update_col_buttons()

    def clear_all_cols(self):
        """清空所有列选择"""
        self.selected_cols.clear()
        self.update_col_buttons()

    def update_selected_cols_display(self):
        """更新已选择列显示"""
        count = len(self.selected_cols)
        if count == 0:
            self.selected_cols_var.set("已选择：0列")
        else:
            # 显示前5个列名
            cols = list(self.selected_cols)[:5]
            display_text = "、".join(cols)
            if count > 5:
                display_text += f" ... 等{count}列"
            else:
                display_text = f"已选择：{display_text}"

            # 限制显示长度
            if len(display_text) > 60:
                display_text = display_text[:57] + "..."

            self.selected_cols_var.set(display_text)

    def add_match_pair(self):
        """添加匹配对"""
        # 创建匹配对框架
        pair_frame = tk.Frame(self.match_pairs_frame, bg=self.colors["card_bg"], pady=5)
        pair_frame.pack(fill=tk.X, pady=3)

        # 主表格列选择
        col1_var = tk.StringVar()
        col1_combo = ttk.Combobox(pair_frame, textvariable=col1_var,
                                  font=self.small_font, state="readonly", width=20)
        col1_combo.pack(side=tk.LEFT, padx=(0, 10))

        # 标签
        tk.Label(pair_frame, text="匹配", font=self.small_font,
                 bg=self.colors["card_bg"]).pack(side=tk.LEFT, padx=(0, 10))

        # 子表格列选择
        col2_var = tk.StringVar()
        col2_combo = ttk.Combobox(pair_frame, textvariable=col2_var,
                                  font=self.small_font, state="readonly", width=20)
        col2_combo.pack(side=tk.LEFT, padx=(0, 10))

        # 匹配规则选择
        rule_var = tk.StringVar(value="fuzzy")
        rule_frame = tk.Frame(pair_frame, bg=self.colors["card_bg"])
        rule_frame.pack(side=tk.LEFT, padx=(0, 10))

        tk.Radiobutton(rule_frame, text="完全相同", variable=rule_var, value="exact",
                       font=self.small_font, bg=self.colors["card_bg"]).pack(side=tk.LEFT)
        tk.Radiobutton(rule_frame, text="模糊匹配", variable=rule_var, value="fuzzy",
                       font=self.small_font, bg=self.colors["card_bg"]).pack(side=tk.LEFT, padx=(5, 0))

        # 删除按钮
        def remove_pair():
            # 从列表中移除
            for i, pair in enumerate(self.match_pairs):
                if pair["frame"] == pair_frame:
                    self.match_pairs.pop(i)
                    break
            # 销毁界面
            pair_frame.destroy()

        del_btn = tk.Button(pair_frame, text="删除", command=remove_pair,
                            font=self.small_font, bg=self.colors["danger"], fg="white",
                            relief=tk.FLAT, padx=5, pady=1)
        del_btn.pack(side=tk.RIGHT)
        self.add_hover_effect(del_btn, self.colors["danger"], "#c92a2a")

        # 保存匹配对信息
        pair_info = {
            "frame": pair_frame,
            "col1_var": col1_var,
            "col1_combo": col1_combo,
            "col2_var": col2_var,
            "col2_combo": col2_combo,
            "rule_var": rule_var
        }

        self.match_pairs.append(pair_info)

        # 更新组合框选项
        self.update_match_pairs_combos()

    def update_match_pairs_combos(self):
        """更新匹配对的组合框选项"""
        # 更新主表格列选项
        if hasattr(self.main_app, 'main_table_data') and self.main_app.main_table_data:
            main_cols = self.main_app.main_table_data["columns"]
            for pair in self.match_pairs:
                pair["col1_combo"]["values"] = main_cols
                if main_cols and not pair["col1_var"].get():
                    pair["col1_var"].set(main_cols[0])

        # 更新子表格列选项
        if self.df is not None:
            sub_cols = list(self.df.columns)
            for pair in self.match_pairs:
                pair["col2_combo"]["values"] = sub_cols
                if sub_cols and not pair["col2_var"].get():
                    pair["col2_var"].set(sub_cols[0])

    def load_existing_data(self):
        """加载现有数据"""
        if not self.existing_data:
            return

        # 加载文件路径
        self.file_path.set(self.existing_data.get("path", ""))
        if self.existing_data.get("path"):
            self.load_file_info()

        # 加载选择的列
        self.selected_cols = set(self.existing_data.get("selected_cols", []))

        # 加载匹配对
        existing_pairs = self.existing_data.get("match_pairs", [])

        # 先清除默认的匹配对
        for pair in self.match_pairs[:]:
            pair["frame"].destroy()
        self.match_pairs.clear()

        # 重新创建匹配对
        for pair_data in existing_pairs:
            self.add_match_pair()

            # 设置匹配对的值
            if self.match_pairs:
                last_pair = self.match_pairs[-1]
                last_pair["col1_var"].set(pair_data.get("col1", ""))
                last_pair["col2_var"].set(pair_data.get("col2", ""))
                last_pair["rule_var"].set(pair_data.get("rule", "fuzzy"))

    def preview_table(self):
        """预览表格"""
        path = self.file_path.get()
        if not path or not os.path.exists(path):
            messagebox.showinfo("提示", "请先选择有效的文件")
            return

        def load_preview():
            try:
                df = self._read_data_file(path, nrows=50)
                self.window.after(0, lambda: PreviewWindow(self.window, df, "子表格预览"))
            except Exception as e:
                self.window.after(0, lambda err=str(e): messagebox.showerror("错误", f"预览失败：{err}"))

        threading.Thread(target=load_preview, daemon=True).start()

    def save_and_close(self):
        """保存设置并关闭窗口"""
        # 验证文件
        path = self.file_path.get()
        if not path or not os.path.exists(path):
            messagebox.showerror("错误", "请选择有效的文件")
            return

        if self.df is None:
            messagebox.showerror("错误", "请等待文件信息加载完成")
            return

        # 验证匹配对
        valid_pairs = []
        for pair in self.match_pairs:
            col1 = pair["col1_var"].get()
            col2 = pair["col2_var"].get()
            rule = pair["rule_var"].get()

            if not col1 or not col2:
                messagebox.showerror("错误", "请为所有匹配对选择列")
                return

            valid_pairs.append({
                "col1": col1,
                "col2": col2,
                "rule": rule
            })

        if not valid_pairs:
            messagebox.showerror("错误", "请至少设置一个匹配条件")
            return

        # 验证选择的列
        if not self.selected_cols:
            messagebox.showerror("错误", "请至少选择一列进行合并")
            return

        # 准备数据
        data = {
            "path": path,
            "df": self.df,
            "full_df": self.full_df,
            "row_count": len(self.full_df),
            "col_count": len(self.df.columns),
            "columns": list(self.df.columns),
            "match_pairs": valid_pairs,
            "selected_cols": list(self.selected_cols)
        }

        # 通知主窗口
        if self.existing_data:
            self.main_app.update_sub_table_data(self.existing_data, data)
        else:
            self.main_app.add_sub_table_data(data)

        # 关闭窗口
        self.window.destroy()

    def on_close(self):
        """关闭窗口"""
        if messagebox.askyesno("确认", "确定取消设置子表格吗？"):
            self.window.destroy()


class PreviewWindow:
    """预览窗口"""

    def __init__(self, parent, df, title="预览"):
        self.parent = parent
        self.df = df

        # 创建窗口
        self.window = tk.Toplevel(parent)
        self.window.title(title)
        self.window.geometry("1000x600")
        self.window.resizable(True, True)
        self.window.transient(parent)

        # 配色
        self.colors = {
            "primary": "#2c6ecb",
            "secondary": "#f5f7fa",
            "card_bg": "#ffffff",
            "border": "#e2e8f0"
        }

        self.font = ("SimHei", 10)
        self.small_font = ("SimHei", 9)

        # 创建界面
        self.create_widgets(title)

    def create_widgets(self, title):
        """创建窗口界面"""
        self.window.configure(bg=self.colors["secondary"])

        # 主容器
        main_container = tk.Frame(self.window, bg=self.colors["secondary"])
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 标题
        title_frame = tk.Frame(main_container, bg=self.colors["primary"], height=40)
        title_frame.pack(fill=tk.X, pady=(0, 10))
        title_frame.pack_propagate(False)

        tk.Label(title_frame, text=title,
                 font=("SimHei", 12, "bold"), bg=self.colors["primary"], fg="white").pack(pady=8)

        # 信息栏
        info_frame = tk.Frame(main_container, bg=self.colors["card_bg"], bd=1, relief=tk.SOLID)
        info_frame.pack(fill=tk.X, pady=(0, 10))

        info_text = f"数据预览：{len(self.df)}行 × {len(self.df.columns)}列"
        if len(self.df) > 50:
            info_text += "（显示前50行）"

        tk.Label(info_frame, text=info_text, font=self.small_font,
                 bg=self.colors["card_bg"], fg="#666666", padx=10, pady=5).pack(anchor=tk.W)

        # 表格容器
        table_container = tk.Frame(main_container, bg=self.colors["secondary"])
        table_container.pack(fill=tk.BOTH, expand=True)

        # 创建Treeview
        self.tree = ttk.Treeview(table_container, show="headings")

        # 设置列
        columns = list(self.df.columns)[:30]  # 最多显示30列
        self.tree["columns"] = columns

        # 配置列
        for col in columns:
            self.tree.heading(col, text=str(col))
            # 自动计算列宽
            col_width = max(80, len(str(col)) * 8)
            self.tree.column(col, width=col_width, anchor=tk.W, minwidth=50)

        # 添加数据
        for i, row in self.df.iterrows():
            values = []
            for col in columns:
                val = row[col]
                if pd.isna(val):
                    values.append("")
                elif isinstance(val, (int, float)):
                    values.append(f"{val:.2f}")
                elif isinstance(val, datetime):
                    values.append(val.strftime("%Y-%m-%d"))
                else:
                    values.append(str(val))
            self.tree.insert("", tk.END, values=values)

        # 添加滚动条
        vscrollbar = ttk.Scrollbar(table_container, orient=tk.VERTICAL, command=self.tree.yview)
        hscrollbar = ttk.Scrollbar(table_container, orient=tk.HORIZONTAL, command=self.tree.xview)

        self.tree.configure(yscrollcommand=vscrollbar.set, xscrollcommand=hscrollbar.set)

        # 布局
        self.tree.grid(row=0, column=0, sticky=tk.NSEW)
        vscrollbar.grid(row=0, column=1, sticky=tk.NS)
        hscrollbar.grid(row=1, column=0, sticky=tk.EW)

        # 配置网格权重
        table_container.grid_rowconfigure(0, weight=1)
        table_container.grid_columnconfigure(0, weight=1)

        # 绑定调整列宽事件
        self.tree.bind("<Button-1>", self.start_resize)
        self.tree.bind("<B1-Motion>", self.on_resize)
        self.resize_column = None
        self.resize_start_x = 0

    def start_resize(self, event):
        """开始调整列宽"""
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            self.resize_column = self.tree.identify_column(event.x)
            self.resize_start_x = event.x

    def on_resize(self, event):
        """调整列宽"""
        if self.resize_column:
            current_width = self.tree.column(self.resize_column, width=None)
            new_width = current_width + (event.x - self.resize_start_x)
            if new_width > 30:
                self.tree.column(self.resize_column, width=new_width)
                self.resize_start_x = event.x


# 主程序入口
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergerPro(root)
    root.mainloop()