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


class ExcelMergerPro:
    def __init__(self, root):
        self.root = root
        self.root.title("数据智能合并工具 - 多表格多条件匹配版")

        # 核心步骤 1: 立即隐藏窗口，防止布局计算过程中的闪烁
        self.root.withdraw()

        # 恢复默认最大化设置 (先不设置，留给后续步骤)
        self.root.minsize(900, 750)  # 最小尺寸限制
        self.root.resizable(True, True)
        self.root.bind("<Configure>=", self.on_window_resize)

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
            "radio_active": "#2c6ecb",
            "radio_inactive": "#cbd5e0",
            "radio_bg": "#ffffff",
            "selected": "#e8f4f8",
            "unselected": "#f8f9fa"
        }

        # 中文字体设置
        self.font = ("SimHei", 10)
        self.small_font = ("SimHei", 9)
        self.radio_font = ("SimHei", 9, "bold")
        self.setup_styles()

        # 数据存储 - 主表格（第一个表格）
        self.file1_path = tk.StringVar()
        self.output_path = tk.StringVar(value="合并结果")
        self.output_format = tk.StringVar(value="xlsx")
        self.df1 = None  # 主表格数据
        self.is_running = False
        self.start_time = 0
        self.total_rows = 0
        self.processed_rows = 0
        self.merge_result_df = None
        self.preview_df = None

        # 子表格列表：每个元素存储一个子表格的完整配置
        # 结构: [{"path_var": StringVar, "path": str, "df": DataFrame, "match_pairs": list, "col_frame": Frame, "selected_cols": set, "info_text": Text, "match_pairs_frame": Frame, "frame": Frame}, ...]
        self.sub_tables = []

        # 线程间通信
        self.progress_queue = Queue()
        self.control_queue = Queue()
        self.log_queue = Queue()
        self.result_queue = Queue()
        self.merge_thread = None

        # Canvas上创建的窗口ID
        self.canvas_window = None

        # 初始化滚动容器
        self.create_scroll_container()

        # 步骤 2: 在隐藏状态下创建组件
        self.create_widgets()

        # 步骤 3: 强制最大化并计算布局
        self.root.state("zoomed")
        self.root.update_idletasks()  # 强制刷新布局，以获得最大化后的正确尺寸
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
        self._on_canvas_resize_for_centering(None)  # 立即执行一次居中更新

        # 步骤 4: 最终显示窗口
        self.root.deiconify()

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
                        thickness=10,
                        troughrelief=tk.FLAT)
        style.configure("TCombobox",
                        fieldbackground=self.colors["card_bg"],
                        background=self.colors["secondary"],
                        foreground=self.colors["text"],
                        borderwidth=1)
        style.map("TCombobox",
                  fieldbackground=[("focus", self.colors["hover"])],
                  bordercolor=[("focus", self.colors["primary_light"])])

    def create_scroll_container(self):
        """创建外层滚动容器（确保右侧滚动条始终可用）"""
        self.root.configure(bg=self.colors["secondary"])

        # 外层容器（铺满整个窗口）
        main_container = tk.Frame(self.root, bg=self.colors["secondary"])
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 滚动容器框架
        scroll_container = tk.Frame(main_container, bg=self.colors["secondary"])
        scroll_container.pack(fill=tk.BOTH, expand=True)

        # 主Canvas（承载所有内容）
        self.main_canvas = tk.Canvas(scroll_container, bg=self.colors["secondary"], highlightthickness=0)

        # 右侧垂直滚动条（始终显示在窗口右侧）
        self.vscrollbar = ttk.Scrollbar(scroll_container, orient="vertical", command=self.main_canvas.yview)
        # 底部水平滚动条
        self.xscrollbar = ttk.Scrollbar(scroll_container, orient="horizontal", command=self.main_canvas.xview)

        # 绑定滚动条和Canvas
        self.main_canvas.configure(
            yscrollcommand=self.vscrollbar.set,
            xscrollcommand=self.xscrollbar.set
        )

        # 滚动内容容器（所有功能组件都放在这里）
        self.scrollable_frame = ttk.Frame(self.main_canvas)
        self.scrollable_frame.configure(style="TFrame")

        # 关键：滚动条滚轮事件强化
        self.vscrollbar.bind("<MouseWheel>", lambda e: self._on_scrollbar_wheel(e, "vertical"))
        self.vscrollbar.bind("<Button-4>", lambda e: self._on_scrollbar_wheel(e, "vertical", delta=120))
        self.vscrollbar.bind("<Button-5>", lambda e: self._on_scrollbar_wheel(e, "vertical", delta=-120))
        self.xscrollbar.bind("<MouseWheel>", lambda e: self._on_scrollbar_wheel(e, "horizontal"))
        self.xscrollbar.bind("<Button-4>", lambda e: self._on_scrollbar_wheel(e, "horizontal", delta=120))
        self.xscrollbar.bind("<Button-5>", lambda e: self._on_scrollbar_wheel(e, "horizontal", delta=-120))
        self.main_canvas.bind("<MouseWheel>", lambda e: self._on_canvas_wheel(e))

        # 在Canvas上创建内容窗口，并保存ID
        self.canvas_window = self.main_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        # 内容变化时更新滚动区域
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
        )

        # 关键：Canvas大小变化时，让内容框架宽度等于Canvas宽度，以实现子组件居中
        self.main_canvas.bind("<Configure>", self._on_canvas_resize_for_centering)

        # 滚动条布局
        self.vscrollbar.pack(side="right", fill="y")
        self.xscrollbar.pack(side="bottom", fill="x")
        self.main_canvas.pack(side="left", fill=tk.BOTH, expand=True)

    def create_widgets(self):
        """在滚动容器中创建所有功能组件"""
        # 顶部标题（居中显示）
        title_font = font.Font(family="SimHei", size=16, weight="bold")
        title_frame = tk.Frame(self.scrollable_frame, bg=self.colors["primary"], height=60)
        title_frame.pack(fill=tk.X, pady=(0, 15))
        title_frame.pack_propagate(False)

        title_label = tk.Label(title_frame,
                               text="数据智能合并工具（多表格联合 | 多条件匹配 | 支持Excel/DTA/CSV）",
                               font=title_font,
                               bg=self.colors["primary"],
                               fg="white")
        title_label.pack(pady=10)

        # 提示文本（居中）
        tip_label = tk.Label(self.scrollable_frame,
                             text="注：1. 主表格为合并基准，子表格将依次合并到主表格 2. 每个子表格的匹配条件需同时满足 3. 支持输入/输出格式：.xlsx .xls .dta .csv",
                             font=("SimHei", 9, "italic"),
                             fg=self.colors["text_light"],
                             bg=self.colors["secondary"])
        tip_label.pack(pady=8)

        # --- 核心内容容器 ---
        self.content_frame = tk.Frame(self.scrollable_frame, bg=self.colors["secondary"])
        self.content_frame.pack(pady=10)

        # 主表格（第一个文件）区域
        main_file_frame = tk.Frame(self.content_frame, bg=self.colors["secondary"])
        main_file_frame.pack(fill=tk.X, pady=8)

        main_file_card = tk.Frame(main_file_frame, bg=self.colors["card_bg"], bd=1, relief=tk.SOLID,
                                  highlightbackground=self.colors["primary"], padx=10, pady=10)
        main_file_card.pack(fill=tk.X)
        main_file_card.grid_columnconfigure(1, weight=3)
        main_file_card.grid_columnconfigure(3, weight=1)

        main_file_shadow = tk.Frame(main_file_frame, bg=self.colors["shadow"], padx=5, pady=5)
        main_file_shadow.pack(fill=tk.X, padx=6, pady=(0, 6))
        main_file_shadow.lower()

        # 主表格标题
        tk.Label(main_file_card, text="📊 主表格（合并基准）：", font=("SimHei", 11, "bold"), bg=self.colors["card_bg"],
                 fg=self.colors["primary"]).grid(
            row=0, column=0, padx=5, pady=(0, 10), sticky=tk.W, columnspan=4)

        # 主表格文件选择
        tk.Label(main_file_card, text="文件路径：", font=self.font, bg=self.colors["card_bg"]).grid(
            row=1, column=0, padx=5, pady=10, sticky=tk.W)
        file1_entry = tk.Entry(main_file_card, textvariable=self.file1_path, font=self.font,
                               bd=1, relief=tk.SUNKEN, bg=self.colors["secondary"])
        file1_entry.grid(row=1, column=1, padx=5, pady=10, sticky=tk.EW)

        # 修正：将 file1_card 改为 main_file_card
        browse1_btn = tk.Button(main_file_card, text="浏览", command=self.browse_main_file, font=self.font,
                                bg=self.colors["primary"], fg="white", relief=tk.FLAT, padx=12, pady=2)
        browse1_btn.grid(row=1, column=2, padx=5, pady=10)
        self.add_hover_effect(browse1_btn, self.colors["primary"], self.colors["primary_light"])

        # 修正：将 file1_card 改为 main_file_card
        preview1_btn = tk.Button(main_file_card, text="预览", command=self.preview_main_file, font=self.small_font,
                                 bg=self.colors["secondary"], fg=self.colors["text"], relief=tk.FLAT, padx=10, pady=2)
        preview1_btn.grid(row=1, column=3, padx=2, pady=10)
        self.add_hover_effect(preview1_btn, self.colors["secondary"], self.colors["hover"])

        # 主表格信息预览
        tk.Label(main_file_card, text="主表格信息：", font=self.font, bg=self.colors["card_bg"]).grid(
            row=2, column=0, padx=5, pady=(10, 5), sticky=tk.W)
        self.main_info_text = tk.Text(main_file_card,
                                      font=self.small_font,
                                      state=tk.DISABLED,
                                      bg=self.colors["secondary"],
                                      bd=1,
                                      relief=tk.SUNKEN,
                                      wrap=tk.WORD,
                                      height=4,
                                      highlightthickness=1,
                                      highlightbackground=self.colors["border"])
        self.main_info_text.grid(row=3, column=0, padx=5, pady=5, sticky=tk.EW, columnspan=4)

        # 输出文件设置（紧跟主表格区域）
        output_frame = tk.Frame(self.content_frame, bg=self.colors["secondary"])
        output_frame.pack(fill=tk.X, pady=8)

        output_card = tk.Frame(output_frame, bg=self.colors["card_bg"], bd=1, relief=tk.SOLID,
                               highlightbackground=self.colors["border"], padx=10, pady=10)
        output_card.pack(fill=tk.X)
        output_card.grid_columnconfigure(1, weight=3)

        tk.Label(output_card, text="💾 输出设置：", font=("SimHei", 11, "bold"), bg=self.colors["card_bg"],
                 fg=self.colors["primary"]).grid(
            row=0, column=0, padx=5, pady=(0, 10), sticky=tk.W, columnspan=5)

        tk.Label(output_card, text="输出文件：", font=self.font, bg=self.colors["card_bg"]).grid(
            row=1, column=0, padx=5, pady=10, sticky=tk.W)
        output_entry = tk.Entry(output_card, textvariable=self.output_path, font=self.font,
                                bd=1, relief=tk.SUNKEN, bg=self.colors["secondary"])
        output_entry.grid(row=1, column=1, padx=5, pady=10, sticky=tk.EW)

        format_label = tk.Label(output_card, text="格式：", font=self.font, bg=self.colors["card_bg"])
        format_label.grid(row=1, column=2, padx=(0, 5), pady=10, sticky=tk.E)

        format_combo = ttk.Combobox(output_card, textvariable=self.output_format,
                                    values=["xlsx", "csv", "dta"], state="readonly",
                                    width=10, font=self.small_font)
        format_combo.grid(row=1, column=3, padx=5, pady=10)
        format_combo.current(0)

        output_btn = tk.Button(output_card, text="选择路径", command=self.browse_output, font=self.font,
                               bg=self.colors["primary"], fg="white", relief=tk.FLAT, padx=12, pady=2)
        output_btn.grid(row=1, column=4, padx=5, pady=10)
        self.add_hover_effect(output_btn, self.colors["primary"], self.colors["primary_light"])

        # 子表格容器（动态添加子表格）
        self.sub_tables_container = tk.Frame(self.content_frame, bg=self.colors["secondary"])
        self.sub_tables_container.pack(fill=tk.X, pady=10)

        # 添加子表格按钮
        add_sub_btn = tk.Button(self.content_frame,
                                text="➕ 添加子表格（可多个）",
                                command=self.add_sub_table,
                                font=("SimHei", 11, "bold"),
                                bg=self.colors["accent"],
                                fg="white",
                                relief=tk.FLAT,
                                padx=20,
                                pady=6)
        add_sub_btn.pack(pady=8)
        self.add_hover_effect(add_sub_btn, self.colors["accent"], "#e64a19")

        # 日志与预览区域
        log_preview_frame = tk.Frame(self.content_frame, bg=self.colors["secondary"])
        log_preview_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # 调整为垂直布局：文件信息预览在上，合并结果预览在下
        log_preview_frame.grid_rowconfigure(0, weight=1)  # 文件信息预览行
        log_preview_frame.grid_rowconfigure(1, weight=2)  # 合并结果预览行（占更大空间）
        log_preview_frame.grid_columnconfigure(0, weight=1)

        # 文件信息预览（横向排列：主表格在左，选中的子表格在右）
        preview_frame = tk.LabelFrame(log_preview_frame,
                                      text="文件信息汇总",
                                      font=self.font,
                                      bg=self.colors["secondary"],
                                      fg=self.colors["text"],
                                      bd=1,
                                      relief=tk.SOLID,
                                      highlightbackground=self.colors["border"],
                                      padx=10,
                                      pady=5)
        preview_frame.grid(row=0, column=0, sticky=tk.NSEW, padx=5, pady=5)
        # 横向分割文件信息预览区域
        preview_frame.grid_columnconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(1, weight=1)
        preview_frame.grid_rowconfigure(0, weight=1)

        # 左侧主表格信息
        self.frame1 = tk.Frame(preview_frame, bg=self.colors["secondary"])
        self.frame1.grid(row=0, column=0, sticky=tk.NSEW, padx=5, pady=5)
        tk.Label(self.frame1,
                 text="主表格信息：",
                 bg=self.colors["secondary"],
                 fg=self.colors["primary"],
                 font=self.font + ("bold",)  # 原有self.font + 加粗属性（元组拼接）
                 ).pack(anchor=tk.W)
        self.info1 = tk.Text(self.frame1,
                             font=self.small_font,
                             state=tk.DISABLED,
                             bg=self.colors["card_bg"],
                             bd=1,
                             relief=tk.SUNKEN,
                             wrap=tk.WORD,
                             highlightthickness=1,
                             highlightbackground=self.colors["border"])
        self.info1.pack(fill=tk.BOTH, expand=True, pady=5)

        # 右侧子表格汇总信息
        self.sub_info_frame = tk.Frame(preview_frame, bg=self.colors["secondary"])
        self.sub_info_frame.grid(row=0, column=1, sticky=tk.NSEW, padx=5, pady=5)
        tk.Label(self.sub_info_frame,
                 text="子表格汇总（共0个）：",
                 bg=self.colors["secondary"],
                 fg=self.colors["accent"],
                 font=("SimHei", 10, "bold")  # 只保留自定义字体，删除重复的 font=self.font
                 ).pack(anchor=tk.W)
        self.sub_info_text = tk.Text(self.sub_info_frame,
                                     font=self.small_font,
                                     state=tk.DISABLED,
                                     bg=self.colors["card_bg"],
                                     bd=1,
                                     relief=tk.SUNKEN,
                                     wrap=tk.WORD,
                                     highlightthickness=1,
                                     highlightbackground=self.colors["border"])
        self.sub_info_text.pack(fill=tk.BOTH, expand=True, pady=5)

        # 合并结果预览（在文件信息预览下方，占更大空间）
        result_frame = tk.LabelFrame(log_preview_frame,
                                     text="合并结果预览（前15条记录）",
                                     font=self.font,
                                     bg=self.colors["secondary"],
                                     fg=self.colors["text"],
                                     bd=1,
                                     relief=tk.SOLID,
                                     highlightbackground=self.colors["border"],
                                     padx=10,
                                     pady=5)
        result_frame.grid(row=1, column=0, sticky=tk.NSEW, padx=5, pady=5)
        result_frame.grid_rowconfigure(0, weight=1)
        result_frame.grid_columnconfigure(0, weight=1)

        # 结果预览表格滚动容器
        tree_scroll_container = tk.Frame(result_frame, bg=self.colors["secondary"])
        tree_scroll_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        tree_scroll_container.grid_columnconfigure(0, weight=1)
        tree_scroll_container.grid_rowconfigure(0, weight=1)

        # 结果预览表格
        self.result_tree = ttk.Treeview(tree_scroll_container, show="headings")
        self.result_tree.grid(row=0, column=0, sticky=tk.NSEW)

        # 垂直滚动条
        vscrollbar = ttk.Scrollbar(tree_scroll_container, orient="vertical", command=self.result_tree.yview)
        vscrollbar.grid(row=0, column=1, sticky=tk.NS)
        self.result_tree.configure(yscrollcommand=vscrollbar.set)

        # 水平滚动条
        hscrollbar = ttk.Scrollbar(result_frame, orient="horizontal", command=self.result_tree.xview)
        hscrollbar.pack(fill=tk.X, padx=5, pady=(0, 5))
        self.result_tree.configure(xscrollcommand=hscrollbar.set)

        # 绑定鼠标滚轮横向滚动
        self.result_tree.bind("<MouseWheel>", self._treeview_horizontal_wheel)

        # 允许拖动调整列宽
        self.result_tree.bind('<Button-1>', self.start_resize)
        self.result_tree.bind('<B1-Motion>', self.on_resize)
        self.resize_column = None
        self.resize_start_x = 0

        # 进度条、状态和时间统计
        progress_frame = tk.Frame(self.content_frame, bg=self.colors["secondary"])
        progress_frame.pack(fill=tk.X, pady=5)
        progress_frame.grid_columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill=tk.X, side=tk.LEFT, padx=5, expand=True)

        self.status_var = tk.StringVar(value="就绪 - 请先选择主表格，再添加子表格")
        status_label = tk.Label(progress_frame, textvariable=self.status_var, font=self.font,
                                bg=self.colors["secondary"])
        status_label.pack(side=tk.LEFT, padx=10)

        self.time_var = tk.StringVar(value="耗时：--:--")
        time_label = tk.Label(progress_frame, textvariable=self.time_var, font=self.font, fg=self.colors["text_light"],
                              bg=self.colors["secondary"])
        time_label.pack(side=tk.RIGHT, padx=10)

        # 操作按钮
        btn_frame = tk.Frame(self.content_frame, bg=self.colors["secondary"], pady=20)
        btn_frame.pack(pady=5)

        button_container = tk.Frame(btn_frame, bg=self.colors["secondary"])
        button_container.pack()

        self.run_btn = tk.Button(button_container, text="开始合并", command=self.start_merge_process,
                                 font=("SimHei", 12, "bold"), width=15,
                                 bg=self.colors["success"], fg="white",
                                 relief=tk.FLAT, padx=5, pady=4)
        self.run_btn.pack(side=tk.LEFT, padx=10)
        self.add_hover_effect(self.run_btn, self.colors["success"], "#4cc964")

        self.cancel_btn = tk.Button(button_container, text="取消", command=self.cancel_merge,
                                    font=("SimHei", 12), width=15,
                                    bg=self.colors["warning"], fg="white",
                                    state=tk.DISABLED, relief=tk.FLAT, padx=5, pady=4)
        self.cancel_btn.pack(side=tk.LEFT, padx=10)
        self.add_hover_effect(self.cancel_btn, self.colors["warning"], "#ffc145")

        clear_btn = tk.Button(button_container, text="清除所有", command=self.clear_selection,
                              font=("SimHei", 12), width=15,
                              bg=self.colors["primary_light"], fg="white",
                              relief=tk.FLAT, padx=5, pady=4)
        clear_btn.pack(side=tk.LEFT, padx=10)
        self.add_hover_effect(clear_btn, self.colors["primary_light"], self.colors["primary"])

        exit_btn = tk.Button(button_container, text="退出",
                             command=lambda: (self.root.destroy(), os._exit(0)),
                             font=("SimHei", 12), width=15,
                             bg=self.colors["danger"], fg="white",
                             relief=tk.FLAT, padx=5, pady=4)
        exit_btn.pack(side=tk.LEFT, padx=10)
        self.add_hover_effect(exit_btn, self.colors["danger"], "#c92a2a")

        # 初始化更新子表格汇总信息
        self.update_sub_tables_summary()

    def add_sub_table(self):
        """动态添加一个子表格配置区域"""
        sub_table_idx = len(self.sub_tables) + 1
        sub_frame = tk.Frame(self.sub_tables_container, bg=self.colors["secondary"], pady=5)
        sub_frame.pack(fill=tk.X, padx=5, pady=5)
        sub_frame.configure(highlightbackground=self.colors["accent"], highlightthickness=2, padx=10, pady=10)

        # 子表格标题栏
        title_frame = tk.Frame(sub_frame, bg=self.colors["secondary"])
        title_frame.pack(fill=tk.X, pady=(0, 10))
        tk.Label(title_frame, text=f"🗂️  子表格 {sub_table_idx}", font=("SimHei", 11, "bold"),
                 bg=self.colors["secondary"], fg=self.colors["accent"]).pack(side=tk.LEFT)

        # 删除按钮
        def remove_sub_table():
            # 从列表中移除
            for i, st in enumerate(self.sub_tables):
                if st["frame"] == sub_frame:
                    self.sub_tables.pop(i)
                    break
            # 销毁界面
            sub_frame.destroy()
            # 更新子表格序号和汇总信息
            self.update_sub_table_titles()
            self.update_sub_tables_summary()

        del_sub_btn = tk.Button(title_frame, text="删除", command=remove_sub_table, font=self.small_font,
                                bg=self.colors["danger"], fg="white", relief=tk.FLAT, padx=8, pady=1)
        del_sub_btn.pack(side=tk.RIGHT)
        self.add_hover_effect(del_sub_btn, self.colors["danger"], "#c92a2a")

        # 文件选择区域
        file_frame = tk.Frame(sub_frame, bg=self.colors["secondary"])
        file_frame.pack(fill=tk.X, pady=5)
        file_frame.grid_columnconfigure(1, weight=3)
        file_frame.grid_columnconfigure(3, weight=1)

        sub_path_var = tk.StringVar()
        tk.Label(file_frame, text="文件路径：", font=self.font, bg=self.colors["secondary"]).grid(
            row=0, column=0, padx=5, pady=5, sticky=tk.W)
        file_entry = tk.Entry(file_frame, textvariable=sub_path_var, font=self.font,
                              bd=1, relief=tk.SUNKEN, bg=self.colors["secondary"])
        file_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)

        # 浏览按钮
        def browse_sub_file():
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
                sub_path_var.set(path)
                # 加载子表格信息
                self.load_sub_table_info(sub_table_idx - 1, path)

        browse_btn = tk.Button(file_frame, text="浏览", command=browse_sub_file, font=self.font,
                               bg=self.colors["primary"], fg="white", relief=tk.FLAT, padx=12, pady=2)
        browse_btn.grid(row=0, column=2, padx=5, pady=5)
        self.add_hover_effect(browse_btn, self.colors["primary"], self.colors["primary_light"])

        # 预览按钮
        def preview_sub_file():
            path = sub_path_var.get()
            if not path or not os.path.exists(path):
                messagebox.showinfo("提示", f"子表格 {sub_table_idx}：请先选择有效的文件")
                return
            self._preview_file(path, f"子表格 {sub_table_idx} 预览")

        preview_btn = tk.Button(file_frame, text="预览", command=preview_sub_file, font=self.small_font,
                                bg=self.colors["secondary"], fg=self.colors["text"], relief=tk.FLAT, padx=10, pady=2)
        preview_btn.grid(row=0, column=3, padx=2, pady=5)
        self.add_hover_effect(preview_btn, self.colors["secondary"], self.colors["hover"])

        # 子表格信息显示
        info_frame = tk.Frame(sub_frame, bg=self.colors["secondary"])
        info_frame.pack(fill=tk.X, pady=5)
        tk.Label(info_frame, text="表格信息：", font=self.font, bg=self.colors["secondary"]).grid(
            row=0, column=0, padx=5, pady=5, sticky=tk.W)
        sub_info_text = tk.Text(info_frame,
                                font=self.small_font,
                                state=tk.DISABLED,
                                bg=self.colors["card_bg"],
                                bd=1,
                                relief=tk.SUNKEN,
                                wrap=tk.WORD,
                                height=3,
                                highlightthickness=1,
                                highlightbackground=self.colors["border"])
        sub_info_text.grid(row=1, column=0, padx=5, pady=5, sticky=tk.EW, columnspan=4)

        # 匹配规则设置区域（LabelFrame，带标题）
        match_frame = tk.LabelFrame(sub_frame,
                                    text=f"匹配规则设置（子表格 {sub_table_idx} - 所有条件必须同时满足）",
                                    font=self.font,
                                    bg=self.colors["secondary"],
                                    fg=self.colors["text"],
                                    bd=1,
                                    relief=tk.SOLID,
                                    highlightbackground=self.colors["border"],
                                    padx=10,
                                    pady=5)
        match_frame.pack(fill=tk.X, pady=8)
        match_frame.grid_columnconfigure(0, weight=1)

        match_pairs_frame = tk.Frame(match_frame, bg=self.colors["secondary"])
        match_pairs_frame.pack(fill=tk.X, padx=5, pady=8)

        # 添加匹配对按钮
        def add_sub_match_pair():
            self.add_sub_match_pair(sub_table_idx - 1)

        add_pair_btn = tk.Button(match_frame,
                                 text="添加匹配对（多条件）",
                                 command=add_sub_match_pair,
                                 font=self.font,
                                 bg=self.colors["primary_light"],
                                 fg="white",
                                 relief=tk.FLAT,
                                 padx=15,
                                 pady=4)
        add_pair_btn.pack(pady=8)
        self.add_hover_effect(add_pair_btn, self.colors["primary_light"], self.colors["primary"])

        # 列选择区域 - LabelFrame（带标题）
        col_select_frame = tk.LabelFrame(sub_frame,
                                         text=f"选择需要合并到主表格的列（子表格 {sub_table_idx}）",
                                         font=self.font,
                                         bg=self.colors["secondary"],
                                         fg=self.colors["text"],
                                         bd=1,
                                         relief=tk.SOLID,
                                         highlightbackground=self.colors["border"],
                                         padx=10,
                                         pady=5)
        col_select_frame.pack(fill=tk.X, pady=8)
        col_select_frame.grid_columnconfigure(0, weight=1)

        # 列选择容器
        col_container = tk.Frame(col_select_frame, bg=self.colors["secondary"])
        col_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 列按钮容器（带滚动条）
        btn_canvas = tk.Canvas(col_container, bg=self.colors["secondary"], highlightthickness=0)
        btn_scrollbar = ttk.Scrollbar(col_container, orient="vertical", command=btn_canvas.yview)
        btn_scrollable_frame = tk.Frame(btn_canvas, bg=self.colors["secondary"])

        btn_scrollable_frame.bind(
            "<Configure>",
            lambda e: btn_canvas.configure(scrollregion=btn_canvas.bbox("all"))
        )

        btn_canvas.create_window((0, 0), window=btn_scrollable_frame, anchor="nw")
        btn_canvas.configure(yscrollcommand=btn_scrollbar.set)

        # 按钮网格布局（每行4个按钮）
        btn_scrollable_frame.grid_columnconfigure((0, 1, 2, 3), weight=1, uniform="col")

        # 已选列显示区域
        selected_cols_frame = tk.Frame(col_select_frame, bg=self.colors["secondary"])
        selected_cols_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        tk.Label(selected_cols_frame, text="已选择列：", font=self.font, bg=self.colors["secondary"]).pack(side=tk.LEFT,
                                                                                                          padx=5)
        self.selected_cols_var = tk.StringVar(value="无")
        selected_cols_label = tk.Label(selected_cols_frame, textvariable=self.selected_cols_var,
                                       font=self.small_font, bg=self.colors["card_bg"],
                                       bd=1, relief=tk.SUNKEN, padx=5, pady=2, wraplength=600)
        selected_cols_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # 按钮布局
        btn_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        btn_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 操作按钮框架
        btn_frame = tk.Frame(col_container, bg=self.colors["secondary"])
        btn_frame.pack(side=tk.RIGHT, padx=10)

        # 全选按钮
        def select_all_sub_cols():
            for widget in btn_scrollable_frame.winfo_children():
                if isinstance(widget, tk.Button):
                    widget.config(bg=self.colors["selected"], relief=tk.SUNKEN)
                    self.sub_tables[sub_table_idx - 1]["selected_cols"].add(widget["text"])
            self.update_selected_cols_display(sub_table_idx - 1)

        select_all_btn = tk.Button(btn_frame, text="全选", command=select_all_sub_cols, font=self.small_font,
                                   bg=self.colors["secondary"], relief=tk.FLAT, padx=10, pady=3)
        select_all_btn.pack(pady=5)
        self.add_hover_effect(select_all_btn, self.colors["secondary"], self.colors["hover"])

        # 取消全选按钮
        def deselect_all_sub_cols():
            for widget in btn_scrollable_frame.winfo_children():
                if isinstance(widget, tk.Button):
                    widget.config(bg=self.colors["unselected"], relief=tk.RAISED)
                    self.sub_tables[sub_table_idx - 1]["selected_cols"].discard(widget["text"])
            self.update_selected_cols_display(sub_table_idx - 1)

        deselect_btn = tk.Button(btn_frame, text="取消全选", command=deselect_all_sub_cols, font=self.small_font,
                                 bg=self.colors["secondary"], relief=tk.FLAT, padx=10, pady=3)
        deselect_btn.pack(pady=5)
        self.add_hover_effect(deselect_btn, self.colors["secondary"], self.colors["hover"])

        # 反选按钮
        def invert_sub_cols():
            for widget in btn_scrollable_frame.winfo_children():
                if isinstance(widget, tk.Button):
                    col_name = widget["text"]
                    if col_name in self.sub_tables[sub_table_idx - 1]["selected_cols"]:
                        widget.config(bg=self.colors["unselected"], relief=tk.RAISED)
                        self.sub_tables[sub_table_idx - 1]["selected_cols"].discard(col_name)
                    else:
                        widget.config(bg=self.colors["selected"], relief=tk.SUNKEN)
                        self.sub_tables[sub_table_idx - 1]["selected_cols"].add(col_name)
            self.update_selected_cols_display(sub_table_idx - 1)

        invert_btn = tk.Button(btn_frame, text="反选", command=invert_sub_cols, font=self.small_font,
                               bg=self.colors["secondary"], relief=tk.FLAT, padx=10, pady=3)
        invert_btn.pack(pady=5)
        self.add_hover_effect(invert_btn, self.colors["secondary"], self.colors["hover"])

        # 保存子表格配置（关键：存入 LabelFrame 引用）
        self.sub_tables.append({
            "path_var": sub_path_var,
            "path": "",
            "df": None,
            "match_pairs": [],
            "match_pairs_frame": match_pairs_frame,  # 普通Frame（承载匹配对）
            "match_label_frame": match_frame,  # LabelFrame（带标题，用于修改文本）
            "col_label_frame": col_select_frame,  # 列选择的LabelFrame（带标题）
            "col_frame": btn_scrollable_frame,  # 列按钮容器
            "selected_cols": set(),  # 存储选中的列名
            "selected_cols_var": self.selected_cols_var,
            "info_text": sub_info_text,
            "frame": sub_frame,
            "add_pair_btn": add_pair_btn,
            "col_container": col_container  # 列选择容器，用于后续销毁
        })

        # 初始添加一对匹配列
        self.add_sub_match_pair(len(self.sub_tables) - 1)
        # 更新汇总信息
        self.update_sub_tables_summary()

    def update_sub_col_buttons(self, sub_table_idx):
        """更新指定子表格的列选择按钮"""
        if sub_table_idx < 0 or sub_table_idx >= len(self.sub_tables):
            return

        sub_table = self.sub_tables[sub_table_idx]
        col_frame = sub_table["col_frame"]
        df = sub_table["df"]

        # 清空现有按钮
        for widget in col_frame.winfo_children():
            widget.destroy()

        if df is None:
            self.update_selected_cols_display(sub_table_idx)
            return

        # 创建列按钮（每行4个）
        cols = list(df.columns)
        for i, col in enumerate(cols):
            # 列按钮点击事件
            def toggle_column(col_name=col):
                if col_name in sub_table["selected_cols"]:
                    sub_table["selected_cols"].discard(col_name)
                    # 找到对应的按钮并更新状态
                    for widget in col_frame.winfo_children():
                        if isinstance(widget, tk.Button) and widget["text"] == col_name:
                            widget.config(bg=self.colors["unselected"], relief=tk.RAISED, fg=self.colors["text"], font=self.small_font)
                            break
                else:
                    sub_table["selected_cols"].add(col_name)
                    # 找到对应的按钮并更新状态
                    for widget in col_frame.winfo_children():
                        if isinstance(widget, tk.Button) and widget["text"] == col_name:
                            widget.config(
                                bg=self.colors["primary"],  # 用主色调作为选中背景（更醒目）
                                fg="black",  # 白色文字
                                font=(self.small_font[0], self.small_font[1], "bold"),  # 文字加粗
                                relief=tk.SUNKEN,
                                bd=2,  # 加粗边框
                                highlightbackground=self.colors["danger"],
                                highlightthickness=2
                            )
                            break
                self.update_selected_cols_display(sub_table_idx)

            # 创建按钮
            btn = tk.Button(col_frame, text=col, command=toggle_column, font=self.small_font,
                            bg=self.colors["unselected"], relief=tk.RAISED, padx=5, pady=3,
                            wraplength=150, justify=tk.CENTER)
            row = i // 4
            col_idx = i % 4
            btn.grid(row=row, column=col_idx, padx=5, pady=3, sticky=tk.EW)
            self.add_hover_effect(btn, self.colors["unselected"], self.colors["hover"])

        self.update_selected_cols_display(sub_table_idx)

    def update_selected_cols_display(self, sub_table_idx):
        """更新已选列显示"""
        if sub_table_idx < 0 or sub_table_idx >= len(self.sub_tables):
            return

        sub_table = self.sub_tables[sub_table_idx]
        selected_cols = sub_table["selected_cols"]

        if not selected_cols:
            sub_table["selected_cols_var"].set("无")
        else:
            # 限制显示长度，超过则截断并显示省略号
            cols_text = "、".join(selected_cols)
            if len(cols_text) > 80:
                cols_text = cols_text[:80] + "..."
            sub_table["selected_cols_var"].set(f"{len(selected_cols)}列：{cols_text}")

    def add_sub_match_pair(self, sub_table_idx):
        """为指定子表格添加匹配对"""
        if sub_table_idx < 0 or sub_table_idx >= len(self.sub_tables):
            return

        sub_table = self.sub_tables[sub_table_idx]
        match_pairs_frame = sub_table["match_pairs_frame"]

        pair_frame = tk.Frame(match_pairs_frame, bg=self.colors["secondary"], pady=3)
        pair_frame.pack(fill=tk.X, padx=5, pady=5)
        pair_frame.configure(highlightbackground=self.colors["border"], highlightthickness=1, padx=5, pady=8)
        # 匹配对中的组件占满宽度
        pair_frame.grid_columnconfigure(0, weight=1)
        pair_frame.grid_columnconfigure(2, weight=1)
        pair_frame.grid_columnconfigure(3, weight=1)

        # 主表格列选择
        col1_var = tk.StringVar()
        col1_combo = ttk.Combobox(pair_frame, textvariable=col1_var, font=self.small_font, state="disabled")
        col1_combo.grid(row=0, column=0, padx=8, sticky=tk.EW)

        # "对"标签
        tk.Label(pair_frame, text="对", font=self.font, bg=self.colors["secondary"]).grid(row=0, column=1, padx=5)

        # 子表格列选择
        col2_var = tk.StringVar()
        col2_combo = ttk.Combobox(pair_frame, textvariable=col2_var, font=self.small_font, state="disabled")
        col2_combo.grid(row=0, column=2, padx=8, sticky=tk.EW)

        # 匹配规则
        rule_var = tk.StringVar(value="fuzzy")
        rule_frame = tk.Frame(pair_frame,
                              bg=self.colors["radio_bg"],
                              bd=1,
                              relief=tk.SOLID,
                              highlightbackground=self.colors["border"],
                              padx=8,
                              pady=2)
        rule_frame.grid(row=0, column=3, padx=10, sticky=tk.EW)

        tk.Label(rule_frame, text="匹配规则：", font=self.small_font, bg=self.colors["radio_bg"],
                 fg=self.colors["text"]).pack(side=tk.LEFT, padx=3)

        exact_radio = tk.Radiobutton(rule_frame,
                                     text="完全相同",
                                     variable=rule_var,
                                     value="exact",
                                     font=self.radio_font,
                                     bg=self.colors["radio_bg"],
                                     fg=self.colors["text"],
                                     activebackground=self.colors["hover"],
                                     indicatoron=1,
                                     width=8)
        exact_radio.pack(side=tk.LEFT, padx=5)

        fuzzy_radio = tk.Radiobutton(rule_frame,
                                     text="模糊匹配",
                                     variable=rule_var,
                                     value="fuzzy",
                                     font=self.radio_font,
                                     bg=self.colors["radio_bg"],
                                     fg=self.colors["text"],
                                     activebackground=self.colors["hover"],
                                     indicatoron=1,
                                     width=8)
        fuzzy_radio.pack(side=tk.LEFT, padx=5)

        # 删除按钮
        def remove_pair():
            for i, pair in enumerate(sub_table["match_pairs"]):
                if pair["frame"] == pair_frame:
                    sub_table["match_pairs"].pop(i)
                    break
            pair_frame.destroy()

        del_btn = tk.Button(pair_frame, text="删除", command=remove_pair, font=self.small_font,
                            bg=self.colors["danger"], fg="white", relief=tk.FLAT, padx=5, pady=1)
        del_btn.grid(row=0, column=4, padx=5)
        self.add_hover_effect(del_btn, self.colors["danger"], "#c92a2a")

        # 添加到子表格的匹配对列表
        sub_table["match_pairs"].append({
            "col1_var": col1_var,
            "col1_combo": col1_combo,
            "col2_var": col2_var,
            "col2_combo": col2_combo,
            "rule_var": rule_var,
            "frame": pair_frame
        })

        # 更新组合框选项（如果主表格和子表格已加载）
        self.update_sub_match_combos(sub_table_idx)

    def update_sub_match_combos(self, sub_table_idx):
        """更新指定子表格的匹配列组合框选项"""
        if sub_table_idx < 0 or sub_table_idx >= len(self.sub_tables):
            return

        sub_table = self.sub_tables[sub_table_idx]

        # 更新主表格列选项
        if self.df1 is not None:
            cols1 = list(self.df1.columns)
            for pair in sub_table["match_pairs"]:
                pair["col1_combo"]["state"] = "readonly"
                pair["col1_combo"]["values"] = cols1
                if cols1 and not pair["col1_var"].get():
                    pair["col1_var"].set(cols1[0])

        # 更新子表格列选项
        if sub_table["df"] is not None:
            cols2 = list(sub_table["df"].columns)
            for pair in sub_table["match_pairs"]:
                pair["col2_combo"]["state"] = "readonly"
                pair["col2_combo"]["values"] = cols2
                if cols2 and not pair["col2_var"].get():
                    pair["col2_var"].set(cols2[0])

    def update_sub_table_titles(self):
        """更新所有子表格的标题序号"""
        for idx, sub_table in enumerate(self.sub_tables):
            sub_idx = idx + 1
            # 更新子表格主标题
            for widget in sub_table["frame"].winfo_children():
                if isinstance(widget, tk.Frame) and widget.winfo_children():
                    first_child = widget.winfo_children()[0]
                    if isinstance(first_child, tk.Label) and first_child["text"].startswith("🗂️  子表格"):
                        first_child["text"] = f"🗂️  子表格 {sub_idx}"
                        break
            # 正确修改匹配规则 LabelFrame 的标题（关键修复）
            sub_table["match_label_frame"]["text"] = f"匹配规则设置（子表格 {sub_idx} - 所有条件必须同时满足）"
            # 正确修改列选择 LabelFrame 的标题（关键修复）
            sub_table["col_label_frame"]["text"] = f"选择需要合并到主表格的列（子表格 {sub_idx}）"

    def update_sub_tables_summary(self):
        """更新子表格汇总信息"""
        summary_text = f"子表格总数：{len(self.sub_tables)}\n"
        summary_text += "已加载的子表格：\n"

        loaded_count = 0
        for idx, sub_table in enumerate(self.sub_tables):
            if sub_table["path"]:
                loaded_count += 1
                filename = os.path.basename(sub_table["path"])
                selected_cols_count = len(sub_table["selected_cols"])
                summary_text += f"  子表格 {idx + 1}：{filename}（{len(sub_table['df'].columns)}列 × {len(sub_table['df'])}行）\n"
                summary_text += f"    - 已选择合并列：{selected_cols_count}列\n"

        if loaded_count == 0:
            summary_text += "  暂无已加载的子表格\n"

        # 更新汇总文本框
        self._update_text(self.sub_info_text, summary_text)
        # 更新状态提示
        if len(self.sub_tables) == 0:
            self.status_var.set("就绪 - 请先选择主表格，再添加子表格")
        else:
            self.status_var.set(f"就绪 - 已添加 {len(self.sub_tables)} 个子表格（{loaded_count} 个已加载）")

    def browse_main_file(self):
        """浏览选择主表格文件"""
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
            self.file1_path.set(path)
            self.load_main_file_info(path)

    def preview_main_file(self):
        """预览主表格文件"""
        self._preview_file(self.file1_path.get(), "主表格预览")

    def load_main_file_info(self, path):
        """加载主表格信息"""

        def load_info():
            try:
                df = self._read_data_file(path, nrows=100)
                self.df1 = df
                info = f"文件名：{os.path.basename(path)}\n"
                info += f"路径：{path}\n"
                info += f"格式：{os.path.splitext(path)[1].upper()}\n"
                info += f"列数：{len(df.columns)}\n"
                total_rows = len(self._read_data_file(path))
                info += f"行数：{total_rows}\n"
                info += "列名（前10列）：\n"
                for i, col in enumerate(df.columns[:10]):
                    info += f"  第{i + 1}列：{col}\n"
                if len(df.columns) > 10:
                    info += f"  ... 共{len(df.columns)}列\n"

                # 更新主表格信息显示
                self.root.after(0, lambda: self._update_text(self.info1, info))
                self.root.after(0, lambda: self._update_text(self.main_info_text, info))

                # 更新所有子表格的匹配组合框（主表格列选项）
                for idx in range(len(self.sub_tables)):
                    self.root.after(0, lambda i=idx: self.update_sub_match_combos(i))

            except Exception as e:
                self.root.after(0, lambda err=str(e): messagebox.showerror("错误", f"读取主表格信息失败：{err}"))

        threading.Thread(target=load_info, daemon=True).start()

    def load_sub_table_info(self, sub_table_idx, path):
        """加载指定子表格信息"""

        def load_info():
            try:
                df = self._read_data_file(path, nrows=100)
                sub_table = self.sub_tables[sub_table_idx]
                sub_table["path"] = path
                sub_table["df"] = df

                info = f"文件名：{os.path.basename(path)}\n"
                info += f"格式：{os.path.splitext(path)[1].upper()}\n"
                info += f"列数：{len(df.columns)} | 行数：{len(self._read_data_file(path))}\n"
                info += "列名（前8列）：\n"
                for i, col in enumerate(df.columns[:8]):
                    info += f"  {col} "
                if len(df.columns) > 8:
                    info += f"... 共{len(df.columns)}列"

                # 更新子表格信息显示
                self.root.after(0, lambda: self._update_text(sub_table["info_text"], info))
                # 更新子表格列选择按钮
                self.root.after(0, lambda: self.update_sub_col_buttons(sub_table_idx))
                # 更新匹配组合框（子表格列选项）
                self.root.after(0, lambda: self.update_sub_match_combos(sub_table_idx))
                # 更新汇总信息
                self.root.after(0, self.update_sub_tables_summary)

            except Exception as e:
                self.root.after(0, lambda err=str(e): messagebox.showerror("错误", f"读取子表格信息失败：{err}"))

        threading.Thread(target=load_info, daemon=True).start()

    # ------------------------------
    # 原有方法修改和完善
    # ------------------------------
    def _read_data_file(self, path, nrows=None):
        ext = os.path.splitext(path)[1].lower()
        try:
            if ext in ['.xlsx', '.xls']:
                df = pd.read_excel(path, nrows=nrows) if nrows else pd.read_excel(path)
            elif ext == '.dta':
                df = pd.read_stata(path)
                df = df.head(nrows) if nrows else df
            elif ext == '.csv':
                encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1']
                for encoding in encodings:
                    try:
                        df = pd.read_csv(path, nrows=nrows, encoding=encoding) if nrows else pd.read_csv(path,
                                                                                                         encoding=encoding)
                        break
                    except:
                        continue
                else:
                    raise ValueError("CSV文件编码无法识别")
            else:
                raise ValueError(f"不支持的文件格式：{ext}")

            # 新增：删除全空行（所有列都为空的行）
            df = df.dropna(how='all')

            return df
        except Exception as e:
            raise Exception(f"读取文件失败：{str(e)}")

    def _update_text(self, text_widget, content):
        text_widget.config(state=tk.NORMAL)
        text_widget.delete(1.0, tk.END)
        text_widget.insert(tk.END, content)
        text_widget.config(state=tk.DISABLED)

    def browse_output(self):
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

    def _preview_file(self, path, title):
        if not path or not os.path.exists(path):
            messagebox.showinfo("提示", "请先选择有效的文件")
            return

        def load_preview():
            try:
                df = self._read_data_file(path, nrows=50)
                self.root.after(0, lambda: self._create_preview_window(df, title))
            except Exception as e:
                self.root.after(0, lambda err=str(e): messagebox.showerror("错误", f"预览失败：{err}"))

        threading.Thread(target=load_preview, daemon=True).start()

    def _create_preview_window(self, df, title):
        preview_window = tk.Toplevel(self.root)
        preview_window.title(title)
        preview_window.state("zoomed")
        preview_window.resizable(True, True)
        preview_window.configure(bg=self.colors["secondary"])

        preview_main_container = tk.Frame(preview_window, bg=self.colors["secondary"])
        preview_main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        preview_scroll_container = tk.Frame(preview_main_container, bg=self.colors["secondary"])
        preview_scroll_container.pack(fill=tk.BOTH, expand=True)

        preview_canvas = tk.Canvas(preview_scroll_container, bg=self.colors["secondary"], highlightthickness=0)
        preview_vscrollbar = ttk.Scrollbar(preview_scroll_container, orient="vertical", command=preview_canvas.yview)
        preview_hscrollbar = ttk.Scrollbar(preview_scroll_container, orient="horizontal", command=preview_canvas.xview)

        preview_canvas.configure(yscrollcommand=preview_vscrollbar.set, xscrollcommand=preview_hscrollbar.set)

        preview_content_frame = ttk.Frame(preview_canvas, style="TFrame")
        preview_content_frame.bind("<Configure>",
                                   lambda e: preview_canvas.configure(scrollregion=preview_canvas.bbox("all")))

        preview_window_id = preview_canvas.create_window((0, 0), window=preview_content_frame, anchor="nw")
        preview_canvas.bind("<Configure>", lambda e: (preview_canvas.itemconfigure(preview_window_id, width=e.width),
                                                      preview_canvas.configure(
                                                          scrollregion=preview_canvas.bbox("all"))))

        preview_vscrollbar.bind("<MouseWheel>", lambda e: preview_canvas.yview_scroll(-int(e.delta / 60), "units"))
        preview_hscrollbar.bind("<MouseWheel>", lambda e: preview_canvas.xview_scroll(-int(e.delta / 60), "units"))
        preview_canvas.bind("<MouseWheel>", lambda e: preview_canvas.yview_scroll(-int(e.delta / 60), "units"))

        preview_vscrollbar.pack(side="right", fill="y")
        preview_hscrollbar.pack(side="bottom", fill="x")
        preview_canvas.pack(side="left", fill=tk.BOTH, expand=True)

        title_frame = tk.Frame(preview_content_frame, bg=self.colors["primary"], height=40)
        title_frame.pack(fill=tk.X, padx=0, pady=0)
        title_frame.pack_propagate(False)
        tk.Label(title_frame, text=title, font=("SimHei", 12, "bold"), bg=self.colors["primary"], fg="white").pack(
            pady=8)

        tree_container = tk.Frame(preview_content_frame, bg=self.colors["secondary"], padx=10, pady=10)
        tree_container.pack(fill=tk.BOTH, expand=True)

        tree = ttk.Treeview(tree_container, show="headings")
        preview_df = self._prepare_preview_df(df)
        columns = list(preview_df.columns[:30])
        tree["columns"] = columns

        width = tree.winfo_width()
        if width > 0 and len(columns) > 0:
            avg_width = max(80, width // len(columns))
            for col in columns:
                tree.heading(col, text=str(col))
                tree.column(col, width=avg_width, anchor=tk.CENTER)
        else:
            for col in columns:
                tree.heading(col, text=str(col))
                tree.column(col, width=100, anchor=tk.CENTER)

        for i, row in preview_df.iterrows():
            values = [str(row[col]) for col in columns]
            tree.insert("", tk.END, values=values)

        vscrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=tree.yview)
        tree.configure(yscroll=vscrollbar.set)
        vscrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        hscrollbar = ttk.Scrollbar(preview_content_frame, orient="horizontal", command=tree.xview)
        tree.configure(xscroll=hscrollbar.set)
        hscrollbar.pack(fill=tk.X, padx=10, pady=(0, 10))

        tree.bind('<Button-1>', self._preview_start_resize)
        tree.bind('<B1-Motion>', lambda e, t=tree: self._preview_on_resize(e, t))
        preview_window.bind("<Configure>", lambda e, t=tree: self._on_preview_window_resize(e, t))

        tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        self._on_preview_window_resize(None, tree)

    def _prepare_preview_df(self, df):
        try:
            preview_df = df.head(15).copy()
            for col in preview_df.columns:
                if pd.api.types.is_datetime64_any_dtype(preview_df[col]):
                    preview_df[col] = preview_df[col].dt.strftime('%Y-%m-%d %H:%M:%S').fillna('')
                elif pd.api.types.is_numeric_dtype(preview_df[col]):
                    preview_df[col] = preview_df[col].round(2).astype(str).replace('nan', '')
                elif pd.api.types.is_bool_dtype(preview_df[col]):
                    preview_df[col] = preview_df[col].map({True: '是', False: '否'}).fillna('')
                else:
                    preview_df[col] = preview_df[col].astype(str).fillna('').str[:50]
            return preview_df
        except Exception as e:
            messagebox.warning("预览警告", f"预览数据处理失败：{str(e)}，将显示原始数据")
            return df.head(15)

    def _display_result_preview(self, df):
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        self.preview_df = self._prepare_preview_df(df)
        max_display_cols = min(25, len(self.preview_df.columns))
        columns = list(self.preview_df.columns[:max_display_cols])
        self.result_tree["columns"] = columns

        for col in columns:
            self.result_tree.heading(col, text=str(col))
            col_width = max(80, len(str(col)) * 8)
            self.result_tree.column(col, width=col_width, anchor=tk.CENTER)

        for i, row in self.preview_df.iterrows():
            values = [str(row[col]) for col in columns]
            self.result_tree.insert("", tk.END, values=values, tags=(i % 2,))

        self.result_tree.tag_configure(0, background=self.colors["card_bg"])
        self.result_tree.tag_configure(1, background=self.colors["hover"])
        self.root.update_idletasks()

    # ------------------------------
    # 合并相关方法修改
    # ------------------------------
    def start_merge_process(self):
        if self.is_running:
            messagebox.showinfo("提示", "合并操作正在进行中，请稍后...")
            return

        # 验证主表格
        main_file = self.file1_path.get()
        if not main_file or not os.path.exists(main_file):
            messagebox.showerror("错误", "请选择有效的主表格文件")
            return
        if self.df1 is None:
            messagebox.showerror("错误", "主表格数据未加载，请重新选择主表格")
            return

        # 验证子表格
        if len(self.sub_tables) == 0:
            messagebox.showerror("错误", "请至少添加一个子表格")
            return

        # 验证每个子表格的配置
        valid_sub_tables = []
        for idx, sub_table in enumerate(self.sub_tables):
            sub_idx = idx + 1
            if not sub_table["path"] or not os.path.exists(sub_table["path"]):
                messagebox.showerror("错误", f"子表格 {sub_idx}：未选择有效的文件")
                return
            if sub_table["df"] is None:
                messagebox.showerror("错误", f"子表格 {sub_idx}：数据未加载，请重新选择")
                return

            # 验证匹配条件
            valid_pairs = []
            for pair_idx, pair in enumerate(sub_table["match_pairs"], 1):
                col1 = pair["col1_var"].get()
                col2 = pair["col2_var"].get()
                rule = pair["rule_var"].get()
                if not col1 or not col2:
                    messagebox.showerror("错误", f"子表格 {sub_idx} - 第{pair_idx}对匹配列未完整选择")
                    return
                valid_pairs.append({"col1": col1, "col2": col2, "rule": rule})
            if not valid_pairs:
                messagebox.showerror("错误", f"子表格 {sub_idx}：请至少添加一对匹配列")
                return

            # 验证选择的列（修改：从set中获取）
            selected_cols = list(sub_table["selected_cols"])
            if not selected_cols:
                messagebox.showerror("错误", f"子表格 {sub_idx}：请选择需要合并的列")
                return

            valid_sub_tables.append({
                "path": sub_table["path"],
                "match_pairs": valid_pairs,
                "selected_cols": selected_cols,
                "index": sub_idx
            })

        # 验证输出路径
        output = self.output_path.get()
        output_format = self.output_format.get()
        if not output.endswith(f".{output_format}"):
            output += f".{output_format}"
            self.output_path.set(output)

        # 准备合并
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        self.is_running = True
        self.start_time = time.time()
        self.processed_rows = 0
        self.total_rows = len(self.df1) * len(valid_sub_tables)  # 总进度基数
        self.run_btn.config(state=tk.DISABLED)
        self.cancel_btn.config(state=tk.NORMAL)
        self.status_var.set(f"开始合并 - 共{len(valid_sub_tables)}个子表格，主表格{len(self.df1)}行数据")
        self.time_var.set("耗时：00:00")
        self.progress["value"] = 0

        # 启动合并线程
        self.merge_thread = threading.Thread(
            target=process_multi_table_merge,
            args=(main_file, valid_sub_tables, output, output_format, self.progress_queue,
                  self.control_queue, self.result_queue),
            daemon=True
        )
        self.merge_thread.start()
        self.root.after(500, self.check_process_status)

    def check_process_status(self):
        if not self.is_running:
            return
        if self.merge_thread.is_alive():
            if not self.control_queue.empty():
                cmd = self.control_queue.get()
                if cmd == "cancel":
                    self.control_queue.put("cancel")
                    self.is_running = False
                    self.status_var.set("合并已取消")
                    self.run_btn.config(state=tk.NORMAL)
                    self.cancel_btn.config(state=tk.DISABLED)
                    return
            self.root.after(500, self.check_process_status)
        else:
            self.is_running = False
            self.progress["value"] = 100
            self.status_var.set("所有表格合并完成！")
            self.run_btn.config(state=tk.NORMAL)
            self.cancel_btn.config(state=tk.DISABLED)
            elapsed = time.time() - self.start_time
            minutes = int(elapsed // 60)
            seconds = int(elapsed % 60)
            self.time_var.set(f"耗时：{minutes:02d}:{seconds:02d}")

            if self.merge_result_df is not None and os.path.exists(self.output_path.get()):
                if messagebox.askyesno("成功", f"所有表格合并完成！文件已保存至：\n{self.output_path.get()}\n是否打开？"):
                    try:
                        os.startfile(self.output_path.get())
                    except:
                        messagebox.showinfo("提示", "文件保存成功，但无法自动打开")
            elif self.merge_result_df is not None:
                messagebox.showwarning("警告", "合并成功，但文件保存失败，可在预览窗口查看结果")
            else:
                messagebox.showwarning("警告", "合并完成，但未获取到结果数据（含有空值）")

    def cancel_merge(self):
        if not self.is_running:
            return
        if messagebox.askyesno("确认", "确定取消所有表格的合并操作？"):
            self.control_queue.put("cancel")
            self.status_var.set("正在取消合并...")
            self.cancel_btn.config(state=tk.DISABLED)

    def clear_selection(self):
        """清除所有选择和配置"""
        # 清除主表格
        self.file1_path.set("")
        self.output_path.set("合并结果")
        self.output_format.set("xlsx")
        self._update_text(self.info1, "")
        self._update_text(self.main_info_text, "")

        # 清除子表格
        for sub_table in self.sub_tables:
            sub_table["frame"].destroy()
        self.sub_tables.clear()

        # 清除结果预览
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        # 清除状态
        self.status_var.set("就绪 - 请先选择主表格，再添加子表格")
        self.time_var.set("耗时：--:--")
        self.progress["value"] = 0
        self.df1 = None
        self.merge_result_df = None
        self.preview_df = None
        self.update_sub_tables_summary()

    # ------------------------------
    # 其他辅助方法（保持不变）
    # ------------------------------
    def _treeview_horizontal_wheel(self, event):
        if event.delta < 0:
            self.result_tree.xview_scroll(1, "units")
        else:
            self.result_tree.xview_scroll(-1, "units")

    def _on_canvas_resize_for_centering(self, event):
        if event is None:
            self.root.update_idletasks()
            canvas_width = self.main_canvas.winfo_width()
        else:
            canvas_width = event.width

        if hasattr(self, 'canvas_window') and canvas_width > 0:
            self.main_canvas.itemconfigure(self.canvas_window, width=canvas_width)

    def on_window_resize(self, event):
        try:
            if hasattr(self, 'result_tree'):
                width = self.result_tree.winfo_width()
                if width > 0 and self.result_tree["columns"]:
                    col_count = len(self.result_tree["columns"])
                    if col_count < 8:
                        avg_width = max(80, width // col_count)
                        for col in self.result_tree["columns"]:
                            self.result_tree.column(col, width=avg_width)
                    else:
                        for col in self.result_tree["columns"]:
                            current_width = self.result_tree.column(col, width=None)
                            if current_width < 80:
                                self.result_tree.column(col, width=80)
        except:
            pass

    def _on_scrollbar_wheel(self, event, direction, delta=None):
        if delta is None:
            delta = event.delta
        scroll_units = -int(delta / 60)
        if direction == "vertical":
            self.main_canvas.yview_scroll(scroll_units, "units")
        else:
            self.main_canvas.xview_scroll(scroll_units, "units")

    def _on_canvas_wheel(self, event):
        delta = event.delta
        scroll_units = -int(delta / 60)
        self.main_canvas.yview_scroll(scroll_units, "units")

    def start_resize(self, event):
        region = self.result_tree.identify_region(event.x, event.y)
        if region == "heading":
            self.resize_column = self.result_tree.identify_column(event.x)
            self.resize_start_x = event.x

    def on_resize(self, event):
        if self.resize_column:
            current_width = self.result_tree.column(self.resize_column, width=None)
            new_width = current_width + (event.x - self.resize_start_x)
            if new_width > 50:
                self.result_tree.column(self.resize_column, width=new_width)
                self.resize_start_x = event.x

    def add_hover_effect(self, button, normal_bg, hover_bg):
        def on_enter(e):
            button['background'] = hover_bg

        def on_leave(e):
            button['background'] = normal_bg

        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)

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
                                progress = int(100 * self.processed_rows / self.total_rows)
                                self.root.after(0, lambda p=progress: self.progress.configure(value=p))
                                self.root.after(0, lambda msg=msg: self.status_var.set(
                                    f"{msg['status']} - 进度：{self.processed_rows}/{self.total_rows}"))
                    elapsed = time.time() - self.start_time
                    minutes = int(elapsed // 60)
                    seconds = int(elapsed % 60)
                    self.root.after(0, lambda m=minutes, s=seconds: self.time_var.set(f"耗时：{m:02d}:{s:02d}"))
                except Exception:
                    break

        threading.Thread(target=listen, daemon=True).start()

    def start_result_listener(self):
        def listen():
            while True:
                try:
                    time.sleep(0.1)
                    if not self.result_queue.empty():
                        result_data = self.result_queue.get()
                        # 强制赋值merge_result_df（无论结果类型）
                        if "df" in result_data:
                            self.merge_result_df = result_data["df"]

                        # 原有逻辑
                        if result_data["type"] == "result":
                            self.root.after(0, lambda df=self.merge_result_df: self._display_result_preview(df))
                        elif result_data["type"] == "error":
                            error_df = pd.DataFrame({"错误信息": [result_data["msg"]]})
                            self.root.after(0, lambda df=error_df: self._display_result_preview(df))
                        elif result_data["type"] == "save_error":
                            self.root.after(0, lambda msg=result_data["msg"]: messagebox.showerror("保存失败", msg))
                            self.root.after(0, lambda df=self.merge_result_df: self._display_result_preview(df))
                        elif result_data["type"] == "warning":
                            self.root.after(0, lambda msg=result_data["msg"]: messagebox.showwarning("提示", msg))
                            self.root.after(0, lambda df=self.merge_result_df: self._display_result_preview(df))
                except Exception as e:
                    print(f"结果监听异常：{str(e)}")
                    break

        threading.Thread(target=listen, daemon=True).start()

    def _on_preview_window_resize(self, event, tree):
        try:
            width = tree.winfo_width()
            columns = tree["columns"]
            if width > 0 and columns:
                col_count = len(columns)
                if col_count < 10:
                    avg_width = max(80, width // col_count)
                    for col in columns:
                        tree.column(col, width=avg_width)
        except:
            pass

    def _preview_start_resize(self, event):
        tree = event.widget
        region = tree.identify_region(event.x, event.y)
        if region == "heading":
            tree.resize_column = tree.identify_column(event.x)
            tree.resize_start_x = event.x

    def _preview_on_resize(self, event, tree):
        if tree.resize_column:
            current_width = tree.column(tree.resize_column, width=None)
            new_width = current_width + (event.x - tree.resize_start_x)
            if new_width > 50:
                tree.column(tree.resize_column, width=new_width)
                tree.resize_start_x = event.x


# ------------------------------
def process_multi_table_merge(main_file_path, sub_tables, output_path, output_format, progress_queue, control_queue,
                              result_queue):
    try:
        # 读取主表格完整数据（已删空行）
        main_df = _read_full_data(main_file_path)
        total_main_rows = len(main_df)
        if total_main_rows == 0:
            raise ValueError("主表格无有效数据（已过滤空行）")

        total_progress_units = total_main_rows * len(sub_tables)
        current_progress = 0

        for sub_table in sub_tables:
            sub_path = sub_table["path"]
            sub_idx = sub_table["index"]
            match_pairs = sub_table["match_pairs"]
            selected_cols = sub_table["selected_cols"]

            # 检查取消信号
            if not control_queue.empty() and control_queue.get() == "cancel":
                result_queue.put(
                    {"type": "error", "msg": f"合并已取消 - 已完成子表格1~{sub_idx - 1}的合并", "df": main_df})
                return

            # 读取子表格并检查有效性
            result_queue.put({"type": "status", "msg": f"正在加载子表格 {sub_idx}..."})
            sub_df = _read_full_data(sub_path)
            if len(sub_df) == 0:
                result_queue.put({"type": "warning", "msg": f"子表格 {sub_idx} 无有效数据（已过滤空行），跳过该子表合并",
                                  "df": main_df})
                continue

            # 准备匹配索引（核心优化：添加匹配数据有效性检查）
            result_queue.put({"type": "status", "msg": f"子表格 {sub_idx} - 准备匹配索引"})
            match_indexes = []
            valid_match_flag = False  # 标记子表是否有可用匹配数据
            for pair_idx, pair in enumerate(match_pairs, 1):
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

                # 构建子表匹配映射（过滤空值）
                index_data = {}
                for idx, val in sub_col.items():
                    if val.strip() != "":
                        index_data.setdefault(val, []).append(idx)
                # 检查子表是否有可用匹配数据
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

            # 子表无可用匹配数据时，跳过合并并提示
            if not valid_match_flag:
                result_queue.put(
                    {"type": "warning", "msg": f"子表格 {sub_idx} 无有效匹配列数据，跳过该子表合并", "df": main_df})
                continue

            # 准备合并列（避免重名）
            target_cols = {}
            for col in selected_cols:
                target_col_name = f"{col}_from_sub{sub_idx}"
                suffix = 1
                while target_col_name in main_df.columns:
                    target_col_name = f"{col}_from_sub{sub_idx}_{suffix}"
                    suffix += 1
                target_cols[col] = target_col_name
                main_df[target_col_name] = None  # 初始化合并列为空

            # 执行合并（优化：增加匹配日志）
            result_queue.put({"type": "status", "msg": f"子表格 {sub_idx} - 开始合并"})
            batch_size = 100
            sub_total_rows = len(main_df)
            matched_count = 0  # 统计匹配成功的行数

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

                        # 精确/模糊匹配逻辑
                        if data["rule"] == "exact":
                            if val1 not in data["index_data"]:
                                all_matched = False
                                break
                            matched_sets.append(set(data["index_data"][val1]))
                        else:
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
                progress_queue.put({
                    "type": "progress",
                    "processed": current_progress,
                    "total": total_progress_units,
                    "status": f"子表格 {sub_idx} 处理中（已匹配{matched_count}行）"
                })

            result_queue.put({"type": "status", "msg": f"子表格 {sub_idx} 合并完成（匹配{matched_count}行）"})

        # 强制确保结果数据传递（核心修复：无论是否匹配成功，都传递主表数据）
        result_queue.put({"type": "result", "df": main_df})

        # 保存结果
        result_queue.put({"type": "status", "msg": "正在保存最终结果..."})
        save_success = _save_output_file(main_df, output_path, output_format)
        if not save_success:
            result_queue.put({"type": "save_error", "msg": "文件保存失败", "df": main_df})

    except Exception as e:
        error_msg = f"合并失败：{str(e)}"
        # 异常时也传递当前主表数据，避免“无结果”
        result_queue.put(
            {"type": "error", "msg": error_msg, "df": main_df if 'main_df' in locals() else pd.DataFrame()})
# ------------------------------
# 数据读取和保存函数（保持不变）
# ------------------------------
def _read_full_data(path):
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

        # 新增：删除全空行（所有列都为空的行）
        df = df.dropna(how='all')

        return df
    except Exception as e:
        raise Exception(f"读取文件失败：{str(e)}")


def _clean_data_for_dta(df):
    df_clean = df.copy()
    df_clean.columns = df_clean.columns.astype(str)
    df_clean.columns = df_clean.columns.str.replace(r'[^\w\s]', '_', regex=True)
    df_clean.columns = df_clean.columns.str.replace(r'\s+', '_', regex=True)
    df_clean.columns = df_clean.columns.str[:32]
    if len(df_clean.columns) != len(set(df_clean.columns)):
        df_clean.columns = [f"col_{i}" for i in range(len(df_clean.columns))]
    for col in df_clean.columns:
        if df_clean[col].dtype == object:
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


def _save_output_file(df, path, output_format):
    try:
        # 先创建数据副本，避免修改原始合并结果
        df_save = df.copy()

        # 统一处理空值（关键修复：针对不同类型列填充对应空值）
        for col in df_save.columns:
            dtype = str(df_save[col].dtype).lower()
            # 字符串/对象类型列：空值填充为空字符串
            if 'object' in dtype or 'string' in dtype:
                df_save[col] = df_save[col].fillna('').astype(str).str[:500]  # 限制长度避免溢出
            # 数值类型列（int/float）：空值填充为 0
            elif 'int' in dtype or 'float' in dtype:
                df_save[col] = df_save[col].replace([np.inf, -np.inf], np.nan).fillna(0)
            # 日期类型列：空值填充为 NaT（pandas 兼容空日期）
            elif 'datetime' in dtype:
                df_save[col] = df_save[col].fillna(pd.NaT)
            # 布尔类型列：空值填充为 False
            elif 'bool' in dtype:
                df_save[col] = df_save[col].fillna(False).astype(int)  # 转为 int 兼容更多格式

        # 原有保存逻辑（基于处理后的干净数据）
        if output_format == "dta":
            df_dta = _clean_data_for_dta(df_save)  # 复用原有 dta 清洗逻辑
            df_dta.to_stata(path, version=114)
        elif output_format == "xlsx":
            df_save.to_excel(path, index=False, engine='openpyxl')
        elif output_format == "csv":
            df_save.to_csv(path, index=False, encoding='utf-8-sig', sep=',', na_rep='')
        else:
            raise ValueError(f"不支持的输出格式：{output_format}")

        return True
    except Exception as e:
        raise Exception(f"保存{output_format.upper()}文件失败：{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergerPro(root)
    root.mainloop()