import tkinter as tk
from tkinter import filedialog, messagebox, ttk, font
import pandas as pd
import os
import threading
from queue import Queue
import time
import math
from typing import List, Optional
import glob

try:
    from ctypes import windll  # Windows DPI适配
except ImportError:
    windll = None


class DtaToCsvConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("DTA转CSV工具 - 分块/直接转换双模式")
        self.root.state("zoomed")
        self.root.geometry("1200x700")  # 加宽主窗口以适应更宽的文件设置区
        self.root.minsize(1000, 600)
        self.root.resizable(True, True)

        # ========== 核心修复：调整执行顺序，先定义colors再初始化字体 ==========
        # 1. 先定义配色方案（避免setup_ttk_styles()找不到colors）
        self.colors = {
            "primary": "#2c3e50",
            "primary_light": "#34495e",
            "secondary": "#f8f9fa",
            "accent": "#3498db",
            "success": "#27ae60",
            "warning": "#f39c12",
            "danger": "#e74c3c",
            "text": "#1a2530",
            "text_light": "#6c757d",
            "card_bg": "#ffffff",
            "border": "#dee2e6"
        }

        # 2. DPI感知设置（与主界面保持一致）
        self.setup_dpi_awareness()

        # 3. 统一字体配置（此时colors已定义，setup_ttk_styles()可正常使用）
        self.setup_unified_fonts()

        # 数据存储（保持不变，新增保留文件结构选项）
        self.source_path = tk.StringVar()
        self.folder_path = tk.StringVar()
        self.output_path = tk.StringVar(value=os.path.join(os.getcwd(), "CSV输出"))
        self.chunk_size = tk.IntVar(value=10000)
        self.encoding = tk.StringVar(value="utf-8-sig")
        self.process_mode = tk.StringVar(value="single")
        self.convert_mode = tk.StringVar(value="direct")
        self.preserve_structure = tk.BooleanVar(value=True)  # 新增：是否保留原始文件结构

        # 状态变量（保持不变）
        self.is_running = False
        self.total_files = 0
        self.processed_files = 0
        self.current_file = ""
        self.current_file_idx = 0
        self.current_chunk = 0
        self.total_chunks = 0

        # 线程通信（保持不变）
        self.progress_queue = Queue()
        self.result_queue = Queue()

        # 创建界面（保持原有布局，优化组件适配）
        self.create_widgets()
        self.start_listeners()

    def setup_dpi_awareness(self):
        """设置DPI感知（与主界面保持一致，避免缩放冲突）"""
        if windll and hasattr(windll.shcore, "SetProcessDpiAwareness"):
            try:
                # 与主界面使用相同的DPI感知级别（1=系统DPI感知）
                windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                pass

    def setup_unified_fonts(self):
        """统一字体配置（与主界面保持一致，解决文字大小差异）"""
        # 字体族：优先使用主界面的SimHei系列，兼容多系统
        font_family = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC", "Microsoft YaHei", "Arial", "sans-serif"]

        # 字体大小：保持14号（与你原代码一致），但统一字体族避免显示差异
        self.font = font.Font(family=font_family, size=11, weight="normal")
        self.title_font = font.Font(family=font_family, size=12, weight="bold")
        self.small_font = font.Font(family=font_family, size=10, weight="normal")  # 微调小字体更清晰

        # 标题栏专用字体（保持12号，统一字体族）
        self.header_font = font.Font(family=font_family, size=12, weight="bold")

        # 同步全局默认字体（关键：避免组件继承主窗口字体冲突）
        default_font = font.nametofont("TkDefaultFont")
        default_font.configure(family=font_family, size=12)

        # 统一ttk组件字体（修复Combobox等控件字体不一致）
        self.setup_ttk_styles()

    def setup_ttk_styles(self):
        """统一ttk组件样式（与主界面协调）"""
        style = ttk.Style()
        # 配置所有ttk组件使用统一字体（此时self.colors已定义，无报错）
        style.configure('TCombobox', font=self.font, background=self.colors["card_bg"], foreground=self.colors["text"])
        style.configure('TProgressbar', troughcolor=self.colors["border"], background=self.colors["accent"])
        style.configure('TLabel', font=self.font)

    def create_widgets(self):
        self.root.configure(bg=self.colors["secondary"])

        # 顶部标题栏（保持原有风格，使用统一字体）
        title_frame = tk.Frame(self.root, bg=self.colors["primary"], height=50)
        title_frame.pack(fill=tk.X, padx=0, pady=0)
        title_frame.pack_propagate(False)

        title_label = tk.Label(
            title_frame,
            text="DTA转CSV工具（分块/直接双模式 | 批量转换 ）",
            font=self.header_font,  # 使用统一标题字体
            bg=self.colors["primary"],
            fg="white"
        )
        title_label.pack(pady=12)

        # 主容器（保持左右布局）
        main_container = tk.Frame(self.root, bg=self.colors["secondary"])
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 左侧窗口：文件设置区（宽度调整为原来的3/2，从480改为720）
        left_frame = tk.LabelFrame(
            main_container,
            text="文件设置",
            font=self.title_font,
            bg=self.colors["secondary"],
            fg=self.colors["text"],
            bd=2,
            relief=tk.SOLID,
            padx=15,
            pady=15
        )
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=5, pady=5)
        left_frame.configure(width=720)  # 原宽度480，改为480*1.5=720
        left_frame.pack_propagate(False)

        # 处理模式选择（保持不变）
        mode_frame = tk.Frame(left_frame, bg=self.colors["secondary"])
        mode_frame.pack(fill=tk.X, pady=5)

        tk.Radiobutton(
            mode_frame,
            text="单个DTA文件",
            variable=self.process_mode,
            value="single",
            font=self.font,
            bg=self.colors["secondary"],
            fg=self.colors["text"],
            command=self.switch_mode,
            selectcolor=self.colors["secondary"]
        ).pack(side=tk.LEFT, padx=10)

        tk.Radiobutton(
            mode_frame,
            text="文件夹批量处理",
            variable=self.process_mode,
            value="folder",
            font=self.font,
            bg=self.colors["secondary"],
            fg=self.colors["text"],
            command=self.switch_mode,
            selectcolor=self.colors["secondary"]
        ).pack(side=tk.LEFT, padx=10)

        # 源文件选择（输入框宽度从55缩短为40，给按钮留出空间）
        self.single_file_frame = tk.Frame(left_frame, bg=self.colors["secondary"])
        self.single_file_frame.pack(fill=tk.X, pady=10)

        tk.Label(
            self.single_file_frame,
            text="源DTA文件：",
            font=self.font,
            bg=self.colors["secondary"],
            fg=self.colors["text"]
        ).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)

        tk.Entry(
            self.single_file_frame,
            textvariable=self.source_path,
            width=40,  # 从55缩短为40，解决按钮遮挡问题
            font=self.font,
            bd=1,
            relief=tk.SUNKEN,
            bg=self.colors["card_bg"],
            fg=self.colors["text"]
        ).grid(row=0, column=1, padx=5, pady=5)

        tk.Button(
            self.single_file_frame,
            text="浏览",
            command=self.select_single_file,
            font=self.font,
            bg=self.colors["accent"],
            fg="white",
            relief=tk.FLAT,
            padx=10
        ).grid(row=0, column=2, padx=5, pady=5)

        # 文件夹选择（默认隐藏，输入框宽度同步缩短为40）
        self.folder_frame = tk.Frame(left_frame, bg=self.colors["secondary"])
        self.folder_frame.pack(fill=tk.X, pady=10)
        self.folder_frame.pack_forget()

        tk.Label(
            self.folder_frame,
            text="源文件夹：",
            font=self.font,
            bg=self.colors["secondary"],
            fg=self.colors["text"]
        ).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)

        tk.Entry(
            self.folder_frame,
            textvariable=self.folder_path,
            width=40,  # 从55缩短为40，解决按钮遮挡问题
            font=self.font,
            bd=1,
            relief=tk.SUNKEN,
            bg=self.colors["card_bg"],
            fg=self.colors["text"]
        ).grid(row=0, column=1, padx=5, pady=5)

        tk.Button(
            self.folder_frame,
            text="选择",
            command=self.select_folder,
            font=self.font,
            bg=self.colors["accent"],
            fg="white",
            relief=tk.FLAT,
            padx=10
        ).grid(row=0, column=2, padx=5, pady=5)

        # 输出路径（输入框宽度同步缩短为40）
        output_frame = tk.Frame(left_frame, bg=self.colors["secondary"])
        output_frame.pack(fill=tk.X, pady=10)

        tk.Label(
            output_frame,
            text="输出路径：",
            font=self.font,
            bg=self.colors["secondary"],
            fg=self.colors["text"]
        ).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)

        tk.Entry(
            output_frame,
            textvariable=self.output_path,
            width=40,  # 从55缩短为40，解决按钮遮挡问题
            font=self.font,
            bd=1,
            relief=tk.SUNKEN,
            bg=self.colors["card_bg"],
            fg=self.colors["text"]
        ).grid(row=0, column=1, padx=5, pady=5)

        tk.Button(
            output_frame,
            text="选择",
            command=self.select_output_folder,
            font=self.font,
            bg=self.colors["accent"],
            fg="white",
            relief=tk.FLAT,
            padx=10
        ).grid(row=0, column=2, padx=5, pady=5)

        # 转换设置（优化网格布局，增加保留文件结构选项）
        settings_frame = tk.LabelFrame(
            left_frame,
            text="转换设置",
            font=self.font,
            bg=self.colors["secondary"],
            padx=10,
            pady=10
        )
        settings_frame.pack(fill=tk.X, pady=15)

        # 转换模式选择（保持优化后的样式）
        convert_mode_frame = tk.Frame(settings_frame, bg=self.colors["secondary"])
        convert_mode_frame.pack(fill=tk.X, pady=8)

        tk.Label(
            convert_mode_frame,
            text="转换模式：",
            font=self.font,
            bg=self.colors["secondary"],
            fg=self.colors["text"]
        ).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)

        self.chunk_radio = tk.Radiobutton(
            convert_mode_frame,
            text="分块处理（大文件）",
            variable=self.convert_mode,
            value="chunk",
            font=self.font,
            bg=self.colors["secondary"],
            fg=self.colors["text"],
            command=self.toggle_chunk_settings,
            selectcolor=self.colors["secondary"],
            activebackground=self.colors["secondary"],
            activeforeground=self.colors["text"]
        ).grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)

        self.direct_radio = tk.Radiobutton(
            convert_mode_frame,
            text="直接转换（小文件）",
            variable=self.convert_mode,
            value="direct",
            font=self.font,
            bg=self.colors["secondary"],
            fg=self.colors["text"],
            command=self.toggle_chunk_settings,
            selectcolor=self.colors["secondary"],
            activebackground=self.colors["secondary"],
            activeforeground=self.colors["text"]
        ).grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)

        # 分块大小（Spinbox宽度同步调整）- 修复创建逻辑
        self.chunk_frame = tk.Frame(settings_frame, bg=self.colors["secondary"])
        self.chunk_frame.pack(fill=tk.X, pady=5)

        tk.Label(
            self.chunk_frame,
            text="分块大小（行）：",
            font=self.font,
            bg=self.colors["secondary"],
            fg=self.colors["text"]
        ).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)

        # 修复：先创建Spinbox对象，再调用grid()
        self.chunk_spinbox = tk.Spinbox(
            self.chunk_frame,
            from_=1000,
            to=100000,
            increment=5000,
            textvariable=self.chunk_size,
            width=15,  # 原宽度12，适配更宽的框架
            font=self.font,
            state=tk.NORMAL,
            bg=self.colors["card_bg"],
            fg=self.colors["text"]
        )
        self.chunk_spinbox.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)

        tk.Label(
            self.chunk_frame,
            text="（大文件建议1-10万行）",
            font=self.small_font,  # 使用优化后的小字体
            bg=self.colors["secondary"],
            fg=self.colors["text_light"]
        ).grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)

        # 编码选择（Combobox宽度同步调整）
        encoding_frame = tk.Frame(settings_frame, bg=self.colors["secondary"])
        encoding_frame.pack(fill=tk.X, pady=5)

        tk.Label(
            encoding_frame,
            text="输出编码：",
            font=self.font,
            bg=self.colors["secondary"],
            fg=self.colors["text"]
        ).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)

        # 复用统一的ttk样式
        encoding_combo = ttk.Combobox(
            encoding_frame,
            textvariable=self.encoding,
            values=["utf-8-sig", "gbk", "gb2312", "utf-8", "latin-1"],
            state="readonly",
            width=17,  # 原宽度14，适配更宽的框架
            style='TCombobox'
        )
        encoding_combo.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        encoding_combo.current(0)

        # 新增：保留原始文件结构选项（批量处理时生效）
        structure_frame = tk.Frame(settings_frame, bg=self.colors["secondary"])
        structure_frame.pack(fill=tk.X, pady=5)

        self.structure_checkbox = tk.Checkbutton(
            structure_frame,
            text="保留原始文件结构（批量处理时生效）",
            variable=self.preserve_structure,
            font=self.font,
            bg=self.colors["secondary"],
            fg=self.colors["text"],
            selectcolor=self.colors["secondary"],
            activebackground=self.colors["secondary"],
            activeforeground=self.colors["text"]
        )
        self.structure_checkbox.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)

        # 右侧窗口：转换状态区（宽度自适应调整）
        right_frame = tk.LabelFrame(
            main_container,
            text="转换状态与结果",
            font=self.title_font,
            bg=self.colors["secondary"],
            fg=self.colors["text"],
            bd=2,
            relief=tk.SOLID,
            padx=15,
            pady=15
        )
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 总进度条（优化长度适配）- 修复标签创建语法
        total_progress_frame = tk.Frame(right_frame, bg=self.colors["secondary"])
        total_progress_frame.pack(fill=tk.X, pady=5)

        tk.Label(
            total_progress_frame,
            text="总进度：",
            font=self.font,
            bg=self.colors["secondary"],
            fg=self.colors["text"]
        ).pack(side=tk.LEFT, padx=5)

        self.total_progress = ttk.Progressbar(
            total_progress_frame,
            orient="horizontal",
            length=250,  # 加长进度条，适配窗口宽度
            mode="determinate"
        )
        self.total_progress.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # 修复：先创建标签对象，再调用pack()
        self.total_progress_label = tk.Label(
            total_progress_frame,
            text="0/0 文件",
            font=self.small_font,
            bg=self.colors["secondary"],
            fg=self.colors["text"]
        )
        self.total_progress_label.pack(side=tk.LEFT, padx=10)

        # 当前文件进度（与总进度条保持一致）- 修复标签创建语法
        file_progress_frame = tk.Frame(right_frame, bg=self.colors["secondary"])
        file_progress_frame.pack(fill=tk.X, pady=5)

        tk.Label(
            file_progress_frame,
            text="当前文件：",
            font=self.font,
            bg=self.colors["secondary"],
            fg=self.colors["text"]
        ).pack(side=tk.LEFT, padx=5)

        self.file_progress = ttk.Progressbar(
            file_progress_frame,
            orient="horizontal",
            length=250,
            mode="determinate"
        )
        self.file_progress.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # 修复：先创建标签对象，再调用pack()
        self.chunk_label = tk.Label(
            file_progress_frame,
            text="分块：0/0",
            font=self.small_font,
            bg=self.colors["secondary"],
            fg=self.colors["text"]
        )
        self.chunk_label.pack(side=tk.LEFT, padx=10)

        # 状态日志文本框（优化高度和滚动体验）
        status_frame = tk.Frame(right_frame, bg=self.colors["secondary"])
        status_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        tk.Label(
            status_frame,
            text="状态日志：",
            font=self.font,
            bg=self.colors["secondary"],
            fg=self.colors["text"]
        ).pack(anchor=tk.W, padx=5)

        self.status_text = tk.Text(
            status_frame,
            height=14,  # 加高日志区域，适配14号字体
            font=self.small_font,
            bg=self.colors["card_bg"],
            bd=1,
            relief=tk.SUNKEN,
            wrap=tk.WORD,
            fg=self.colors["text"],
            insertbackground=self.colors["text"]
        )
        self.status_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.status_text.config(state=tk.DISABLED)

        # 添加垂直滚动条（优化日志浏览）
        scrollbar = ttk.Scrollbar(self.status_text, orient=tk.VERTICAL, command=self.status_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.config(yscrollcommand=scrollbar.set)

        # 操作按钮区域（保持原有样式）
        btn_frame = tk.Frame(self.root, bg=self.colors["secondary"])
        btn_frame.pack(fill=tk.X, pady=10)

        self.start_btn = tk.Button(
            btn_frame,
            text="开始转换",
            command=self.start_conversion,
            font=self.title_font,
            bg=self.colors["success"],
            fg="white",
            relief=tk.FLAT,
            padx=30,
            pady=5
        )
        self.start_btn.pack(side=tk.LEFT, padx=20)

        self.stop_btn = tk.Button(
            btn_frame,
            text="停止转换",
            command=self.stop_conversion,
            font=self.title_font,
            bg=self.colors["danger"],
            fg="white",
            relief=tk.FLAT,
            padx=30,
            pady=5,
            state=tk.DISABLED
        )
        self.stop_btn.pack(side=tk.LEFT, padx=20)

        self.open_btn = tk.Button(
            btn_frame,
            text="打开输出文件夹",
            command=self.open_output_folder,
            font=self.title_font,
            bg=self.colors["accent"],
            fg="white",
            relief=tk.FLAT,
            padx=30,
            pady=5
        )
        self.open_btn.pack(side=tk.LEFT, padx=20)

        # 初始化切换模式
        self.switch_mode()
        self.toggle_chunk_settings()

        # 强制刷新界面，确保样式生效
        self.root.update_idletasks()

    # ========== 以下修改convert_files方法，添加保留文件结构逻辑 ==========
    def switch_mode(self):
        if self.process_mode.get() == "single":
            self.single_file_frame.pack(fill=tk.X, pady=10)
            self.folder_frame.pack_forget()
            self.structure_checkbox.config(state=tk.DISABLED, fg=self.colors["text_light"])
        else:
            self.single_file_frame.pack_forget()
            self.folder_frame.pack(fill=tk.X, pady=10)
            self.structure_checkbox.config(state=tk.NORMAL, fg=self.colors["text"])

    def toggle_chunk_settings(self):
        if self.convert_mode.get() == "chunk":
            self.chunk_spinbox.config(state=tk.NORMAL)
            for widget in self.chunk_frame.winfo_children():
                if widget.winfo_class() == "Label":
                    widget.config(fg=self.colors["text"])
        else:
            self.chunk_spinbox.config(state=tk.DISABLED)
            for widget in self.chunk_frame.winfo_children():
                if widget.winfo_class() == "Label":
                    widget.config(fg=self.colors["text_light"])

    def select_single_file(self):
        file_path = filedialog.askopenfilename(
            title="选择DTA文件",
            filetypes=[("DTA文件", "*.dta"), ("所有文件", "*.*")]
        )
        if file_path:
            self.source_path.set(file_path)
            self.log_status(f"已选择单个文件：{file_path}")

    def select_folder(self):
        folder_path = filedialog.askdirectory(title="选择包含DTA文件的文件夹")
        if folder_path:
            self.folder_path.set(folder_path)
            self.log_status(f"已选择文件夹：{folder_path}")

    def select_output_folder(self):
        folder_path = filedialog.askdirectory(title="选择CSV输出文件夹")
        if folder_path:
            self.output_path.set(folder_path)
            self.log_status(f"已选择输出路径：{folder_path}")

    def open_output_folder(self):
        output_path = self.output_path.get()
        if os.path.exists(output_path):
            try:
                os.startfile(output_path)
            except:
                messagebox.showinfo("提示", "无法自动打开文件夹，请手动访问：\n" + output_path)
        else:
            messagebox.showwarning("警告", "输出文件夹不存在！")

    def log_status(self, message: str):
        self.status_text.config(state=tk.NORMAL)
        timestamp = time.strftime("%H:%M:%S", time.localtime())
        self.status_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.status_text.see(tk.END)
        self.status_text.config(state=tk.DISABLED)

    def start_listeners(self):
        def progress_listener():
            while True:
                try:
                    time.sleep(0.1)
                    if not self.progress_queue.empty():
                        data = self.progress_queue.get()
                        self.root.after(0, self.update_progress, data)
                except:
                    break

        def result_listener():
            while True:
                try:
                    time.sleep(0.1)
                    if not self.result_queue.empty():
                        result = self.result_queue.get()
                        self.root.after(0, self.handle_result, result)
                except:
                    break

        threading.Thread(target=progress_listener, daemon=True).start()
        threading.Thread(target=result_listener, daemon=True).start()

    def update_progress(self, data: dict):
        if data["type"] == "file_start":
            self.current_file = data["filename"]
            self.current_chunk = 0
            self.total_chunks = data.get("total_chunks", 0)
            self.current_file_idx = data.get("current_idx", 1)
            self.total_files = data.get("total_files", 1)

            if self.convert_mode.get() == "direct":
                self.file_progress["value"] = 100
                self.chunk_label.config(text="直接转换中...")
            else:
                self.file_progress["value"] = 0
                self.chunk_label.config(text=f"分块：0/{self.total_chunks}")

            self.log_status(
                f"开始转换 [{self.current_file_idx}/{self.total_files}]：{os.path.basename(self.current_file)}")

        elif data["type"] == "update_chunks":
            self.total_chunks = data["total_chunks"]
            self.chunk_label.config(text=f"分块：0/{self.total_chunks}")

        elif data["type"] == "chunk_done":
            self.current_chunk += 1
            chunk_progress = (self.current_chunk / self.total_chunks) * 100
            self.file_progress["value"] = chunk_progress
            self.chunk_label.config(text=f"分块：{self.current_chunk}/{self.total_chunks}")
            self.log_status(f"完成分块 {self.current_chunk}/{self.total_chunks}")

        elif data["type"] == "file_done":
            self.processed_files += 1
            current_idx = data.get("current_idx", self.processed_files)
            total_files = data.get("total_files", self.total_files)
            total_progress = (self.processed_files / total_files) * 100
            self.total_progress["value"] = total_progress
            self.total_progress_label.config(text=f"{self.processed_files}/{total_files} 文件")
            self.log_status(f"完成转换 [{current_idx}/{total_files}]：{os.path.basename(self.current_file)}")

        elif data["type"] == "total_files":
            self.total_files = data["count"]
            self.total_progress_label.config(text=f"0/{self.total_files} 文件")

        elif data["type"] == "conversion_complete":
            self.log_status("=" * 50)
            self.log_status(f"转换完成！共处理 {self.processed_files}/{self.total_files} 个文件")
            self.log_status(f"输出路径：{self.output_path.get()}")
            self.log_status(f"转换模式：{'直接转换' if self.convert_mode.get() == 'direct' else '分块处理'}")
            self.log_status(
                f"保留文件结构：{'是' if self.preserve_structure.get() and self.process_mode.get() == 'folder' else '否'}")
            self.log_status("=" * 50)

        elif data["type"] == "scan_file":
            self.log_status(f"扫描到DTA文件：{data['file_path']}")

    def handle_result(self, result: dict):
        total_files = result.get("total_files", self.total_files)
        total_processed = result.get("total_processed", self.processed_files)

        if result["type"] == "success":
            pass
        elif result["type"] == "error":
            self.log_status(f"错误：{result['message']}")
            if "批量转换失败" in result["message"]:
                messagebox.showerror("错误", result["message"])
        elif result["type"] == "stopped":
            self.log_status("=" * 50)
            self.log_status(f"转换已停止！已处理 {total_processed}/{total_files} 个文件")
            self.log_status(
                f"保留文件结构：{'是' if self.preserve_structure.get() and self.process_mode.get() == 'folder' else '否'}")
            self.log_status("=" * 50)
            messagebox.showinfo("提示", f"转换已停止！\n已处理 {total_processed}/{total_files} 个文件")

        self.is_running = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)

    def validate_input(self) -> tuple[bool, str]:
        if self.process_mode.get() == "single":
            if not self.source_path.get() or not os.path.exists(self.source_path.get()):
                return False, "请选择有效的DTA文件！"
            if not self.source_path.get().lower().endswith(".dta"):
                return False, "选择的文件不是DTA格式！"
        else:
            if not self.folder_path.get() or not os.path.exists(self.folder_path.get()):
                return False, "请选择有效的文件夹！"

        if self.convert_mode.get() == "chunk":
            if self.chunk_size.get() < 1000 or self.chunk_size.get() > 100000:
                return False, "分块大小必须在1000-100000行之间！"

        output_path = self.output_path.get()
        if not os.path.exists(output_path):
            try:
                os.makedirs(output_path)
                self.log_status(f"已创建输出文件夹：{output_path}")
            except:
                return False, f"无法创建输出文件夹：{output_path}"

        return True, ""

    def get_dta_files(self) -> List[str]:
        dta_files = []
        if self.process_mode.get() == "single":
            if self.source_path.get() and os.path.exists(self.source_path.get()):
                dta_files.append(self.source_path.get())
                self.progress_queue.put({"type": "scan_file", "file_path": self.source_path.get()})
        else:
            folder_path = self.folder_path.get()
            if folder_path and os.path.exists(folder_path):
                self.log_status(f"开始扫描文件夹：{folder_path}（包含子文件夹）")
                for root, dirs, files in os.walk(folder_path):
                    self.log_status(f"正在扫描子文件夹：{root}")
                    for file in files:
                        if file.lower().endswith(".dta"):
                            full_path = os.path.abspath(os.path.join(root, file))
                            dta_files.append(full_path)
                            self.progress_queue.put({"type": "scan_file", "file_path": full_path})

        # 优化：移除不必要的去重，直接排序（保留原始顺序的同时排序）
        dta_files.sort()
        return dta_files

    def start_conversion(self):
        valid, msg = self.validate_input()
        if not valid:
            messagebox.showerror("输入错误", msg)
            return

        self.is_running = True
        self.processed_files = 0
        self.current_file_idx = 0
        self.total_files = 0
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)

        self.status_text.config(state=tk.NORMAL)
        self.status_text.delete(1.0, tk.END)
        self.status_text.config(state=tk.DISABLED)

        self.log_status("开始查找DTA文件...")
        dta_files = self.get_dta_files()
        actual_total = len(dta_files)

        if actual_total == 0:
            self.log_status("未找到任何.dta格式的文件！")
            messagebox.showwarning("警告", "未找到任何DTA文件！\n请检查选择的路径是否正确，或文件是否为.dta格式。")
            self.is_running = False
            self.start_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            return

        self.log_status(f"\n共找到 {actual_total} 个DTA文件，待处理清单：")
        for i, file in enumerate(dta_files, 1):
            self.log_status(f"  {i}. {os.path.basename(file)}")
        self.log_status("\n开始转换...")
        if self.process_mode.get() == "folder" and self.preserve_structure.get():
            self.log_status("提示：已启用保留原始文件结构模式")

        self.progress_queue.put({"type": "total_files", "count": actual_total})

        conversion_thread = threading.Thread(
            target=self.convert_files,
            args=(dta_files, self.chunk_size.get(), self.encoding.get(), self.output_path.get()),
            daemon=False
        )
        conversion_thread.start()

    def stop_conversion(self):
        if messagebox.askyesno("确认", "确定要停止转换吗？"):
            self.is_running = False
            self.stop_btn.config(state=tk.DISABLED)
            self.log_status("正在停止转换...")

    def convert_files(self, dta_files: List[str], chunk_size: int, encoding: str, output_path: str):
        actual_total = len(dta_files)
        source_folder = self.folder_path.get() if self.process_mode.get() == "folder" else ""

        try:
            for idx, dta_file in enumerate(dta_files):
                if not self.is_running:
                    break

                current_file_idx = idx + 1
                filename = os.path.splitext(os.path.basename(dta_file))[0]

                # 处理输出路径：如果保留文件结构，则创建对应的子文件夹
                if self.process_mode.get() == "folder" and self.preserve_structure.get():
                    # 获取文件相对于源文件夹的相对路径
                    relative_path = os.path.relpath(os.path.dirname(dta_file), source_folder)
                    # 构建输出文件夹路径
                    output_file_folder = os.path.join(output_path, relative_path)
                    # 创建文件夹（如果不存在）
                    if not os.path.exists(output_file_folder):
                        os.makedirs(output_file_folder)
                    # 构建最终的CSV路径
                    csv_file = os.path.join(output_file_folder, f"{filename}.csv")
                else:
                    # 不保留文件结构，直接输出到根目录
                    csv_file = os.path.join(output_path, f"{filename}.csv")

                if os.path.exists(csv_file):
                    self.log_status(f"警告：{os.path.basename(csv_file)} 已存在，将覆盖")

                try:
                    self.progress_queue.put({
                        "type": "file_start",
                        "filename": dta_file,
                        "total_chunks": 1 if self.convert_mode.get() == "direct" else 0,
                        "current_idx": current_file_idx,
                        "total_files": actual_total
                    })

                    if self.convert_mode.get() == "direct":
                        self.log_status(f"直接转换中（小文件优化）...")
                        df = pd.read_stata(dta_file)
                        df = self.clean_data(df)
                        df.to_csv(
                            csv_file,
                            index=False,
                            encoding=encoding,
                            mode='w',
                            header=True
                        )
                    else:
                        total_rows = self.get_dta_row_count(dta_file)
                        total_chunks = math.ceil(total_rows / chunk_size) if total_rows > 0 else 1

                        self.progress_queue.put({
                            "type": "update_chunks",
                            "total_chunks": total_chunks
                        })

                        first_chunk = True
                        for chunk_idx, chunk in enumerate(pd.read_stata(dta_file, chunksize=chunk_size)):
                            if not self.is_running:
                                break

                            chunk = self.clean_data(chunk)
                            chunk.to_csv(
                                csv_file,
                                index=False,
                                encoding=encoding,
                                mode='w' if first_chunk else 'a',
                                header=first_chunk
                            )

                            first_chunk = False
                            self.progress_queue.put({"type": "chunk_done"})

                    if self.is_running:
                        self.progress_queue.put({
                            "type": "file_done",
                            "current_idx": current_file_idx,
                            "total_files": actual_total
                        })
                        self.log_status(f"输出路径：{csv_file}")

                except Exception as e:
                    error_msg = f"转换 [{current_file_idx}/{actual_total}] {os.path.basename(dta_file)} 失败：{str(e)}"
                    self.log_status(error_msg)
                    continue

            if self.is_running:
                self.progress_queue.put({"type": "conversion_complete"})
                self.result_queue.put({
                    "type": "success",
                    "total_processed": self.processed_files,
                    "total_files": actual_total
                })
            else:
                self.result_queue.put({
                    "type": "stopped",
                    "total_processed": self.processed_files,
                    "total_files": actual_total
                })

        except Exception as e:
            error_msg = f"批量转换失败：{str(e)}"
            self.log_status(error_msg)
            self.result_queue.put({
                "type": "error",
                "message": error_msg,
                "total_files": actual_total
            })

    def get_dta_row_count(self, dta_file: str) -> int:
        try:
            df_sample = pd.read_stata(dta_file, chunksize=1)
            total_rows = 0
            for chunk in df_sample:
                total_rows += len(chunk)
            return total_rows
        except:
            df = pd.read_stata(dta_file)
            return len(df)

    def clean_data(self, df: pd.DataFrame) -> pd.DataFrame:
        df_clean = df.copy()

        df_clean.columns = df_clean.columns.astype(str)
        df_clean.columns = df_clean.columns.str.replace(r'[^\w\s]', '_', regex=True)
        df_clean.columns = df_clean.columns.str.replace(r'\s+', '_', regex=True)

        for col in df_clean.columns:
            if df_clean[col].dtype == 'object':
                df_clean[col] = df_clean[col].fillna('')
                df_clean[col] = df_clean[col].astype(str).str.encode('utf-8', errors='replace').str.decode('utf-8')
            elif 'datetime64' in str(df_clean[col].dtype):
                df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce').dt.strftime('%Y-%m-%d %H:%M:%S')
            df_clean[col] = df_clean[col].fillna('')

        return df_clean


if __name__ == "__main__":
    root = tk.Tk()
    app = DtaToCsvConverter(root)
    root.mainloop()