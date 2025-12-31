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


class ExcelMergerPro_2:
    def __init__(self, root):
        self.root = root
        self.root.title("数据智能合并工具 - 多条件匹配版")

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
            "radio_bg": "#ffffff"
        }

        # 中文字体设置
        self.font = ("SimHei", 10)
        self.small_font = ("SimHei", 9)
        self.radio_font = ("SimHei", 9, "bold")
        self.setup_styles()

        # 数据存储
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.output_path = tk.StringVar(value="合并结果")
        self.output_format = tk.StringVar(value="xlsx")
        self.df1 = None
        self.df2 = None
        self.selected_cols = []
        self.is_running = False
        self.start_time = 0
        self.total_rows = 0
        self.processed_rows = 0
        self.merge_result_df = None
        self.preview_df = None

        # 线程间通信
        self.progress_queue = Queue()
        self.control_queue = Queue()
        self.log_queue = Queue()
        self.result_queue = Queue()
        self.merge_thread = None

        # 匹配配置
        self.match_pairs: List[Dict] = []

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
        # 标题栏横跨整个窗口（仍保持铺满，这不影响内容区不铺满）
        title_frame = tk.Frame(self.scrollable_frame, bg=self.colors["primary"], height=60)
        title_frame.pack(fill=tk.X, pady=(0, 15))
        title_frame.pack_propagate(False)

        title_label = tk.Label(title_frame,
                               text="数据智能合并工具（多条件联合匹配 | 支持Excel/DTA/CSV）",
                               font=title_font,
                               bg=self.colors["primary"],
                               fg="white")
        title_label.pack(pady=10)

        # 提示文本（居中）
        tip_label = tk.Label(self.scrollable_frame,
                             text="注：所有匹配对必须同时匹配成功才会执行合并 | 支持输入/输出格式：.xlsx .xls .dta .csv",
                             font=("SimHei", 9, "italic"),
                             fg=self.colors["text_light"],
                             bg=self.colors["secondary"])
        tip_label.pack(pady=8)

        # --- 核心内容容器 (取消铺满，仅根据内容大小居中) ---
        self.content_frame = tk.Frame(self.scrollable_frame, bg=self.colors["secondary"])
        # 不使用 fill=tk.X 或 expand=True，组件将根据内部内容决定宽度，并在父容器中默认居中。
        self.content_frame.pack(pady=10)
        # ------------------------------------

        # 文件选择区域（现在以 self.content_frame 为父容器）
        file_frame = tk.Frame(self.content_frame, bg=self.colors["secondary"])
        file_frame.pack(fill=tk.X, pady=8)

        file_card = tk.Frame(file_frame, bg=self.colors["card_bg"], bd=1, relief=tk.SOLID,
                             highlightbackground=self.colors["border"], padx=10, pady=10)
        file_card.pack(fill=tk.X)
        # 输入框列占主要宽度
        file_card.grid_columnconfigure(1, weight=3)
        file_card.grid_columnconfigure(3, weight=1)

        file_shadow = tk.Frame(file_frame, bg=self.colors["shadow"], padx=5, pady=5)
        file_shadow.pack(fill=tk.X, padx=6, pady=(0, 6))
        file_shadow.lower()

        # 第一个文件
        tk.Label(file_card, text="第一个数据文件：", font=self.font, bg=self.colors["card_bg"]).grid(
            row=0, column=0, padx=5, pady=10, sticky=tk.W)
        file1_entry = tk.Entry(file_card, textvariable=self.file1_path, font=self.font,
                               bd=1, relief=tk.SUNKEN, bg=self.colors["secondary"])
        file1_entry.grid(row=0, column=1, padx=5, pady=10, sticky=tk.EW)

        browse1_btn = tk.Button(file_card, text="浏览", command=self.browse_file1, font=self.font,
                                bg=self.colors["primary"], fg="white", relief=tk.FLAT, padx=12, pady=2)
        browse1_btn.grid(row=0, column=2, padx=5, pady=10)
        self.add_hover_effect(browse1_btn, self.colors["primary"], self.colors["primary_light"])

        preview1_btn = tk.Button(file_card, text="预览", command=self.preview_file1, font=self.small_font,
                                 bg=self.colors["secondary"], fg=self.colors["text"], relief=tk.FLAT, padx=10, pady=2)
        preview1_btn.grid(row=0, column=3, padx=2, pady=10)
        self.add_hover_effect(preview1_btn, self.colors["secondary"], self.colors["hover"])

        # 第二个文件
        tk.Label(file_card, text="第二个数据文件：", font=self.font, bg=self.colors["card_bg"]).grid(
            row=1, column=0, padx=5, pady=10, sticky=tk.W)
        file2_entry = tk.Entry(file_card, textvariable=self.file2_path, font=self.font,
                               bd=1, relief=tk.SUNKEN, bg=self.colors["secondary"])
        file2_entry.grid(row=1, column=1, padx=5, pady=10, sticky=tk.EW)

        browse2_btn = tk.Button(file_card, text="浏览", command=self.browse_file2, font=self.font,
                                bg=self.colors["primary"], fg="white", relief=tk.FLAT, padx=12, pady=2)
        browse2_btn.grid(row=1, column=2, padx=5, pady=10)
        self.add_hover_effect(browse2_btn, self.colors["primary"], self.colors["primary_light"])

        preview2_btn = tk.Button(file_card, text="预览", command=self.preview_file2, font=self.small_font,
                                 bg=self.colors["secondary"], fg=self.colors["text"], relief=tk.FLAT, padx=10, pady=2)
        preview2_btn.grid(row=1, column=3, padx=2, pady=10)
        self.add_hover_effect(preview2_btn, self.colors["secondary"], self.colors["hover"])

        # 输出文件 + 格式选择
        tk.Label(file_card, text="输出文件：", font=self.font, bg=self.colors["card_bg"]).grid(
            row=2, column=0, padx=5, pady=10, sticky=tk.W)
        output_entry = tk.Entry(file_card, textvariable=self.output_path, font=self.font,
                                bd=1, relief=tk.SUNKEN, bg=self.colors["secondary"])
        output_entry.grid(row=2, column=1, padx=5, pady=10, sticky=tk.EW)

        format_label = tk.Label(file_card, text="格式：", font=self.font, bg=self.colors["card_bg"])
        format_label.grid(row=2, column=2, padx=(0, 5), pady=10, sticky=tk.E)

        format_combo = ttk.Combobox(file_card, textvariable=self.output_format,
                                    values=["xlsx", "csv", "dta"], state="readonly",
                                    width=10, font=self.small_font)
        format_combo.grid(row=2, column=3, padx=5, pady=10)
        format_combo.current(0)

        output_btn = tk.Button(file_card, text="选择路径", command=self.browse_output, font=self.font,
                               bg=self.colors["primary"], fg="white", relief=tk.FLAT, padx=12, pady=2)
        output_btn.grid(row=2, column=4, padx=5, pady=10)
        self.add_hover_effect(output_btn, self.colors["primary"], self.colors["primary_light"])

        # 匹配设置区域 (现在以 self.content_frame 为父容器)
        match_frame = tk.LabelFrame(self.content_frame,
                                    text="匹配规则设置（所有条件必须同时满足）",
                                    font=self.font,
                                    bg=self.colors["secondary"],
                                    fg=self.colors["text"],
                                    bd=1,
                                    relief=tk.SOLID,
                                    highlightbackground=self.colors["border"],
                                    padx=10,
                                    pady=5)
        match_frame.pack(fill=tk.X, pady=10)
        match_frame.grid_columnconfigure(0, weight=1)

        self.match_pairs_frame = tk.Frame(match_frame, bg=self.colors["secondary"])
        self.match_pairs_frame.pack(fill=tk.X, padx=5, pady=8)

        add_pair_btn = tk.Button(match_frame,
                                 text="添加匹配对（多条件）",
                                 command=self.add_match_pair,
                                 font=self.font,
                                 bg=self.colors["primary_light"],
                                 fg="white",
                                 relief=tk.FLAT,
                                 padx=15,
                                 pady=4)
        add_pair_btn.pack(pady=8)
        self.add_hover_effect(add_pair_btn, self.colors["primary_light"], self.colors["primary"])

        # 列选择区域 (现在以 self.content_frame 为父容器)
        col_select_frame = tk.LabelFrame(self.content_frame,
                                         text="选择第二个表中需要合并的列",
                                         font=self.font,
                                         bg=self.colors["secondary"],
                                         fg=self.colors["text"],
                                         bd=1,
                                         relief=tk.SOLID,
                                         highlightbackground=self.colors["border"],
                                         padx=10,
                                         pady=5)
        col_select_frame.pack(fill=tk.X, pady=10)
        col_select_frame.grid_columnconfigure(0, weight=1)

        # 列选择容器
        col_container = tk.Frame(col_select_frame, bg=self.colors["secondary"])
        col_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 列选择列表框
        self.col_listbox = tk.Listbox(col_container,
                                      selectmode=tk.MULTIPLE,
                                      font=self.small_font,
                                      height=6,
                                      bg=self.colors["card_bg"],
                                      bd=1,
                                      relief=tk.SUNKEN,
                                      selectbackground=self.colors["primary_light"],
                                      selectforeground="white",
                                      activestyle="none",
                                      highlightthickness=1,
                                      highlightbackground=self.colors["border"],
                                      highlightcolor=self.colors["primary"])
        self.col_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10)

        # 列表框滚动条
        col_vscrollbar = ttk.Scrollbar(col_container, orient="vertical", command=self.col_listbox.yview)
        col_hscrollbar = ttk.Scrollbar(col_select_frame, orient="horizontal", command=self.col_listbox.xview)
        self.col_listbox.configure(yscrollcommand=col_vscrollbar.set, xscrollcommand=col_hscrollbar.set)
        col_vscrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        col_hscrollbar.pack(fill=tk.X, padx=10, pady=(0, 10))

        btn_frame = tk.Frame(col_container, bg=self.colors["secondary"])
        btn_frame.pack(side=tk.RIGHT, padx=10)

        select_all_btn = tk.Button(btn_frame, text="全选", command=self.select_all_cols, font=self.small_font,
                                   bg=self.colors["secondary"], relief=tk.FLAT, padx=10, pady=3)
        select_all_btn.pack(pady=5)
        self.add_hover_effect(select_all_btn, self.colors["secondary"], self.colors["hover"])

        deselect_btn = tk.Button(btn_frame, text="取消全选", command=self.deselect_all_cols, font=self.small_font,
                                 bg=self.colors["secondary"], relief=tk.FLAT, padx=10, pady=3)
        deselect_btn.pack(pady=5)
        self.add_hover_effect(deselect_btn, self.colors["secondary"], self.colors["hover"])

        # 日志与预览区域 (现在以 self.content_frame 为父容器)
        log_preview_frame = tk.Frame(self.content_frame, bg=self.colors["secondary"])
        log_preview_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # 调整为垂直布局：文件信息预览在上，合并结果预览在下
        log_preview_frame.grid_rowconfigure(0, weight=1)  # 文件信息预览行
        log_preview_frame.grid_rowconfigure(1, weight=2)  # 合并结果预览行（占更大空间）
        log_preview_frame.grid_columnconfigure(0, weight=1)

        # 文件信息预览（横向排列：文件1在左，文件2在右）
        preview_frame = tk.LabelFrame(log_preview_frame,
                                      text="文件信息预览",
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

        # 左侧文件1信息
        self.frame1 = tk.Frame(preview_frame, bg=self.colors["secondary"])
        self.frame1.grid(row=0, column=0, sticky=tk.NSEW, padx=5, pady=5)
        tk.Label(self.frame1, text="文件1信息：", font=self.font, bg=self.colors["secondary"],
                 fg=self.colors["text"]).pack(anchor=tk.W)
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

        # 右侧文件2信息
        self.frame2 = tk.Frame(preview_frame, bg=self.colors["secondary"])
        self.frame2.grid(row=0, column=1, sticky=tk.NSEW, padx=5, pady=5)
        tk.Label(self.frame2, text="文件2信息：", font=self.font, bg=self.colors["secondary"],
                 fg=self.colors["text"]).pack(anchor=tk.W)
        self.info2 = tk.Text(self.frame2,
                             font=self.small_font,
                             state=tk.DISABLED,
                             bg=self.colors["card_bg"],
                             bd=1,
                             relief=tk.SUNKEN,
                             wrap=tk.WORD,
                             highlightthickness=1,
                             highlightbackground=self.colors["border"])
        self.info2.pack(fill=tk.BOTH, expand=True, pady=5)

        # 合并结果预览（在文件信息预览下方，占更大空间）- 修复水平滚动条
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
        result_frame.grid_columnconfigure(0, weight=1)  # 确保容器宽度自适应

        # 关键修复1：创建独立的滚动容器，包含Treeview和垂直滚动条
        tree_scroll_container = tk.Frame(result_frame, bg=self.colors["secondary"])
        tree_scroll_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        tree_scroll_container.grid_columnconfigure(0, weight=1)
        tree_scroll_container.grid_rowconfigure(0, weight=1)

        # 结果预览表格
        self.result_tree = ttk.Treeview(tree_scroll_container, show="headings")
        self.result_tree.grid(row=0, column=0, sticky=tk.NSEW)  # 使用grid布局确保自适应

        # 垂直滚动条
        vscrollbar = ttk.Scrollbar(tree_scroll_container, orient="vertical", command=self.result_tree.yview)
        vscrollbar.grid(row=0, column=1, sticky=tk.NS)
        self.result_tree.configure(yscrollcommand=vscrollbar.set)

        # 关键修复2：水平滚动条放在result_frame，绑定Treeview的xview
        hscrollbar = ttk.Scrollbar(result_frame, orient="horizontal", command=self.result_tree.xview)
        hscrollbar.pack(fill=tk.X, padx=5, pady=(0, 5))  # 放在Treeview容器下方
        self.result_tree.configure(xscrollcommand=hscrollbar.set)

        # 关键修复3：绑定鼠标滚轮横向滚动
        self.result_tree.bind("<MouseWheel>", self._treeview_horizontal_wheel)

        # 允许拖动调整列宽
        self.result_tree.bind('<Button-1>', self.start_resize)
        self.result_tree.bind('<B1-Motion>', self.on_resize)
        self.resize_column = None
        self.resize_start_x = 0

        # 进度条、状态和时间统计 (现在以 self.content_frame 为父容器)
        progress_frame = tk.Frame(self.content_frame, bg=self.colors["secondary"])
        progress_frame.pack(fill=tk.X, pady=5)
        progress_frame.grid_columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill=tk.X, side=tk.LEFT, padx=5, expand=True)

        self.status_var = tk.StringVar(value="就绪")
        status_label = tk.Label(progress_frame, textvariable=self.status_var, font=self.font,
                                bg=self.colors["secondary"])
        status_label.pack(side=tk.LEFT, padx=10)

        self.time_var = tk.StringVar(value="耗时：--:--")
        time_label = tk.Label(progress_frame, textvariable=self.time_var, font=self.font, fg=self.colors["text_light"],
                              bg=self.colors["secondary"])
        time_label.pack(side=tk.RIGHT, padx=10)

        # 操作按钮 (现在以 self.content_frame 为父容器)
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

        clear_btn = tk.Button(button_container, text="清除选择", command=self.clear_selection,
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

        # 初始添加一对匹配列
        self.add_match_pair()

        # 这一部分布局更新和居中逻辑已被移动到 __init__ 的末尾，以确保在 deiconify() 前执行。

    def _treeview_horizontal_wheel(self, event):
        """Treeview鼠标滚轮横向滚动支持"""
        if event.delta < 0:
            self.result_tree.xview_scroll(1, "units")
        else:
            self.result_tree.xview_scroll(-1, "units")

    def _on_canvas_resize_for_centering(self, event):
        """
        Canvas大小变化时，确保scrollable_frame的宽度与其相等。
        这样scrollable_frame内的子组件（如self.content_frame）
        才能使用pack()实现居中。
        """
        if event is None:
            # 首次加载时，需要先更新root和canvas的实际尺寸
            self.root.update_idletasks()
            canvas_width = self.main_canvas.winfo_width()
        else:
            canvas_width = event.width

        # 更新Canvas上窗口的宽度
        if hasattr(self, 'canvas_window') and canvas_width > 0:
            self.main_canvas.itemconfigure(self.canvas_window, width=canvas_width)

    def on_window_resize(self, event):
        """窗口大小变化时更新表格列宽"""
        try:
            # 调整表格列宽
            if hasattr(self, 'result_tree'):
                width = self.result_tree.winfo_width()
                if width > 0 and self.result_tree["columns"]:
                    col_count = len(self.result_tree["columns"])
                    # 关键修复：不强制平均分配列宽，保留原始列宽以触发横向滚动
                    if col_count > 0:
                        # 仅在列数较少时调整，列数多时保留原始宽度
                        if col_count < 8:
                            avg_width = max(80, width // col_count)
                            for col in self.result_tree["columns"]:
                                self.result_tree.column(col, width=avg_width)
                        else:
                            # 列数多时，固定列宽触发横向滚动
                            for col in self.result_tree["columns"]:
                                current_width = self.result_tree.column(col, width=None)
                                if current_width < 80:
                                    self.result_tree.column(col, width=80)
        except:
            pass

    # 滚动条滚轮控制核心方法
    def _on_scrollbar_wheel(self, event, direction, delta=None):
        """滚动条上的滚轮控制"""
        if delta is None:
            delta = event.delta
        # 滚动单位（控制滚动速度）
        scroll_units = -int(delta / 60)  # 加快滚动速度
        if direction == "vertical":
            self.main_canvas.yview_scroll(scroll_units, "units")
        else:
            self.main_canvas.xview_scroll(scroll_units, "units")

    def _on_canvas_wheel(self, event):
        """Canvas内容区的滚轮控制"""
        delta = event.delta
        scroll_units = -int(delta / 60)
        # 垂直滚动优先
        self.main_canvas.yview_scroll(scroll_units, "units")

    # 列宽调整功能
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

    def add_match_pair(self):
        """动态添加匹配对"""
        pair_frame = tk.Frame(self.match_pairs_frame, bg=self.colors["secondary"], pady=3)
        pair_frame.pack(fill=tk.X, padx=5, pady=5)
        pair_frame.configure(highlightbackground=self.colors["border"], highlightthickness=1, padx=5, pady=8)
        # 匹配对中的组件占满宽度
        pair_frame.grid_columnconfigure(0, weight=1)
        pair_frame.grid_columnconfigure(2, weight=1)
        pair_frame.grid_columnconfigure(3, weight=1)

        # 表1列选择
        col1_var = tk.StringVar()
        col1_combo = ttk.Combobox(pair_frame, textvariable=col1_var, font=self.small_font, state="disabled")
        col1_combo.grid(row=0, column=0, padx=8, sticky=tk.EW)

        # "对"标签
        tk.Label(pair_frame, text="对", font=self.font, bg=self.colors["secondary"]).grid(row=0, column=1, padx=5)

        # 表2列选择
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
            for i, pair in enumerate(self.match_pairs):
                if pair["frame"] == pair_frame:
                    self.match_pairs.pop(i)
                    break
            pair_frame.destroy()

        del_btn = tk.Button(pair_frame, text="删除", command=remove_pair, font=self.small_font,
                            bg=self.colors["danger"], fg="white", relief=tk.FLAT, padx=5, pady=1)
        del_btn.grid(row=0, column=4, padx=5)
        self.add_hover_effect(del_btn, self.colors["danger"], "#c92a2a")

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
            self.load_file_info(path, 1)

    def browse_file2(self):
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
            self.file2_path.set(path)
            self.load_file_info(path, 2)

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

    def _read_data_file(self, path, nrows=None):
        ext = os.path.splitext(path)[1].lower()
        try:
            if ext in ['.xlsx', '.xls']:
                return pd.read_excel(path, nrows=nrows) if nrows else pd.read_excel(path)
            elif ext == '.dta':
                df = pd.read_stata(path)
                return df.head(nrows) if nrows else df
            elif ext == '.csv':
                encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1']
                for encoding in encodings:
                    try:
                        return pd.read_csv(path, nrows=nrows, encoding=encoding) if nrows else pd.read_csv(path,
                                                                                                           encoding=encoding)
                    except:
                        continue
                raise ValueError("CSV文件编码无法识别")
            else:
                raise ValueError(f"不支持的文件格式：{ext}")
        except Exception as e:
            raise Exception(f"读取文件失败：{str(e)}")

    def load_file_info(self, path, file_num):
        def load_info():
            try:
                df = self._read_data_file(path, nrows=100)
                info = f"文件名：{os.path.basename(path)}\n"
                info += f"路径：{path}\n"
                info += f"格式：{os.path.splitext(path)[1].upper()}\n"
                info += f"列数：{len(df.columns)}\n"
                total_rows = len(self._read_data_file(path))
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

        # 关键修复：设置固定列宽，确保内容超出时触发横向滚动
        for col in columns:
            self.result_tree.heading(col, text=str(col))
            # 根据列名长度设置基础宽度，避免过窄
            col_width = max(80, len(str(col)) * 8)
            self.result_tree.column(col, width=col_width, anchor=tk.CENTER)

        # 插入数据
        for i, row in self.preview_df.iterrows():
            values = [str(row[col]) for col in columns]
            self.result_tree.insert("", tk.END, values=values, tags=(i % 2,))

        # 交替行颜色
        self.result_tree.tag_configure(0, background=self.colors["card_bg"])
        self.result_tree.tag_configure(1, background=self.colors["hover"])

        # 强制刷新滚动区域
        self.root.update_idletasks()

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
                df = self._read_data_file(path, nrows=50)
                self.root.after(0, lambda: self._create_preview_window(df, title))
            except Exception as e:
                self.root.after(0, lambda err=str(e): messagebox.showerror("错误", f"预览失败：{err}"))

        threading.Thread(target=load_preview, daemon=True).start()

    def _create_preview_window(self, df, title):
        preview_window = tk.Toplevel(self.root)
        preview_window.title(title)
        preview_window.state("zoomed")  # 预览窗口也最大化
        preview_window.resizable(True, True)
        preview_window.configure(bg=self.colors["secondary"])

        # 预览窗口滚动容器
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

        # 预览窗口滚动条滚轮控制
        preview_vscrollbar.bind("<MouseWheel>", lambda e: preview_canvas.yview_scroll(-int(e.delta / 60), "units"))
        preview_hscrollbar.bind("<MouseWheel>", lambda e: preview_canvas.xview_scroll(-int(e.delta / 60), "units"))
        preview_canvas.bind("<MouseWheel>", lambda e: preview_canvas.yview_scroll(-int(e.delta / 60), "units"))

        preview_vscrollbar.pack(side="right", fill="y")
        preview_hscrollbar.pack(side="bottom", fill="x")
        preview_canvas.pack(side="left", fill=tk.BOTH, expand=True)

        # 预览窗口标题
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

        # 初始时强制更新列宽
        self._on_preview_window_resize(None, tree)

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

    def clear_selection(self):
        self.file1_path.set("")
        self.file2_path.set("")
        self.output_path.set("合并结果")
        self.output_format.set("xlsx")
        self._update_text(self.info1, "")
        self._update_text(self.info2, "")
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
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
        self.merge_result_df = None
        self.preview_df = None

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
                                    f"处理第{self.processed_rows}/{self.total_rows}行..."))
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
                        if result_data["type"] == "result":
                            self.merge_result_df = result_data["df"]
                            self.root.after(0, lambda df=self.merge_result_df: self._display_result_preview(df))
                        elif result_data["type"] == "error":
                            error_df = pd.DataFrame({"错误信息": [result_data["msg"]]})
                            self.root.after(0, lambda df=error_df: self._display_result_preview(df))
                        elif result_data["type"] == "save_error":
                            self.root.after(0, lambda msg=result_data["msg"]: messagebox.showerror("保存失败", msg))
                            if result_data.get("df") is not None:
                                self.root.after(0, lambda df=result_data["df"]: self._display_result_preview(df))
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
        output_format = self.output_format.get()

        if not output.endswith(f".{output_format}"):
            output += f".{output_format}"
            self.output_path.set(output)

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

        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        self.is_running = True
        self.start_time = time.time()
        self.processed_rows = 0
        self.run_btn.config(state=tk.DISABLED)
        self.cancel_btn.config(state=tk.NORMAL)
        self.status_var.set("正在初始化处理线程...")
        self.time_var.set("耗时：00:00")
        self.progress["value"] = 5

        self.merge_thread = threading.Thread(
            target=process_merge,
            args=(file1, file2, output, output_format, valid_pairs, self.selected_cols, self.progress_queue,
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
            self.status_var.set("合并完成！")
            self.run_btn.config(state=tk.NORMAL)
            self.cancel_btn.config(state=tk.DISABLED)
            elapsed = time.time() - self.start_time
            minutes = int(elapsed // 60)
            seconds = int(elapsed % 60)
            self.time_var.set(f"耗时：{minutes:02d}:{seconds:02d}")
            if self.merge_result_df is not None and os.path.exists(self.output_path.get()):
                if messagebox.askyesno("成功", f"文件已保存至：\n{self.output_path.get()}\n是否打开？"):
                    try:
                        os.startfile(self.output_path.get())
                    except:
                        messagebox.showinfo("提示", "文件保存成功，但无法自动打开")
            elif self.merge_result_df is not None:
                messagebox.showwarning("警告", "合并成功，但文件保存失败，可在预览窗口查看结果")
            else:
                messagebox.showwarning("警告", "合并完成，但未获取到结果数据")

    def cancel_merge(self):
        if not self.is_running:
            return
        if messagebox.askyesno("确认", "确定取消？"):
            self.control_queue.put("cancel")
            self.status_var.set("正在取消...")
            self.cancel_btn.config(state=tk.DISABLED)


def _read_full_data(path):
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext in ['.xlsx', '.xls']:
            return pd.read_excel(path)
        elif ext == '.dta':
            return pd.read_stata(path)
        elif ext == '.csv':
            encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1']
            for encoding in encodings:
                try:
                    return pd.read_csv(path, encoding=encoding)
                except:
                    continue
            raise ValueError("CSV文件编码无法识别")
        else:
            raise ValueError(f"不支持的文件格式：{ext}")
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
            # 修复警告：添加 infer_objects(copy=False) 替代隐式向下转换
            df_clean[col] = df_clean[col].fillna('').infer_objects(copy=False).astype(str).str.encode('utf-8', errors='replace').str.decode(
                'latin-1').str[:244]
        elif 'datetime64' in str(df_clean[col].dtype):
            df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce').dt.date
        elif str(df_clean[col].dtype).startswith('category'):
            # 修复警告：添加 infer_objects(copy=False)
            df_clean[col] = df_clean[col].astype(str).fillna('').infer_objects(copy=False).str.encode('utf-8', errors='replace').str.decode(
                'latin-1')
        elif df_clean[col].dtype == bool:
            df_clean[col] = df_clean[col].astype(int)
        elif 'int' in str(df_clean[col].dtype).lower() or 'float' in str(df_clean[col].dtype).lower():
            df_clean[col] = df_clean[col].replace([np.inf, -np.inf], np.nan).fillna(0)
    return df_clean


def _save_output_file(df, path, output_format):
    try:
        if output_format == "dta":
            df_dta = _clean_data_for_dta(df)
            df_dta.to_stata(path, version=114)
        elif output_format == "xlsx":
            df.to_excel(path, index=False, engine='openpyxl')
        elif output_format == "csv":
            df.to_csv(path, index=False, encoding='utf-8-sig', sep=',', na_rep='')
        else:
            raise ValueError(f"不支持的输出格式：{output_format}")
        return True
    except Exception as e:
        raise Exception(f"保存{output_format.upper()}文件失败：{str(e)}")


def process_merge(file1, file2, output, output_format, match_pairs, selected_cols, progress_queue, control_queue,
                  result_queue):
    try:
        df1 = _read_full_data(file1)
        df2 = _read_full_data(file2)
        total_rows = len(df1)
        if total_rows == 0 or len(df2) == 0:
            raise ValueError("文件无数据")

        result_queue.put({"type": "status", "msg": f"开始多条件联合匹配（共{len(match_pairs)}个匹配条件）"})
        match_indexes = []
        for pair_idx, pair in enumerate(match_pairs, 1):
            col1, col2, rule = pair["col1"], pair["col2"], pair["rule"]
            if col1 not in df1.columns:
                raise ValueError(f"匹配对{pair_idx}：文件1中不存在列 '{col1}'")
            if col2 not in df2.columns:
                raise ValueError(f"匹配对{pair_idx}：文件2中不存在列 '{col2}'")

            df1_col = df1[col1].fillna("").astype(str).str.strip()
            df2_col = df2[col2].fillna("").astype(str).str.strip()
            if rule == "fuzzy":
                df1_col = df1_col.str.lower()
                df2_col = df2_col.str.lower()

            index_data = {}
            for idx, val in df2_col.items():
                if val not in index_data:
                    index_data[val] = []
                index_data[val].append(idx)
            match_indexes.append({
                "pair_idx": pair_idx,
                "col1": col1,
                "col2": col2,
                "rule": rule,
                "df1_col": df1_col,
                "index_data": index_data,
                "df2": df2
            })

        target_cols = {col: f"{col}_from_file2" for col in selected_cols}
        for col in target_cols.values():
            if col not in df1.columns:
                df1[col] = None

        result_queue.put({"type": "status", "msg": f"开始合并数据，共{total_rows}行待处理"})
        batch_size = 100
        for i in range(0, total_rows, batch_size):
            # 检查是否需要取消
            if not control_queue.empty():
                if control_queue.get() == "cancel":
                    result_queue.put({"type": "error", "msg": "合并已取消"})
                    return

            end = min(i + batch_size, total_rows)
            for idx in range(i, end):
                all_matched = True
                matched_sets = []
                for data in match_indexes:
                    val1 = data["df1_col"].iloc[idx]
                    if not val1:
                        all_matched = False
                        break
                    if data["rule"] == "exact":
                        if val1 not in data["index_data"]:
                            all_matched = False
                            break
                        matched_sets.append(set(data["index_data"][val1]))
                    else:  # 模糊匹配
                        matched_ids = []
                        for val2, ids in data["index_data"].items():
                            if val1 in val2 or val2 in val1:
                                matched_ids.extend(ids)
                        if not matched_ids:
                            all_matched = False
                            break
                        matched_sets.append(set(matched_ids))

                if all_matched and matched_sets:
                    # 求所有匹配集的交集
                    common_ids = set.intersection(*matched_sets)
                    if common_ids:
                        # 取第一个匹配项的数据
                        match_id = next(iter(common_ids))
                        for col, target_col in target_cols.items():
                            df1.at[idx, target_col] = df2.at[match_id, col]

            # 更新进度
            progress_queue.put({"type": "progress", "processed": end, "total": total_rows})

        # 保存文件
        result_queue.put({"type": "status", "msg": "正在保存结果文件..."})
        save_success = _save_output_file(df1, output, output_format)
        if save_success:
            result_queue.put({"type": "result", "df": df1})
        else:
            result_queue.put({"type": "save_error", "msg": "文件保存失败", "df": df1})

    except Exception as e:
        error_msg = f"合并失败：{str(e)}"
        result_queue.put({"type": "error", "msg": error_msg})


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergerPro_2(root)
    root.mainloop()