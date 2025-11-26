import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, font
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os
import matplotlib
from datetime import datetime
from sklearn.decomposition import PCA
from sklearn.preprocessing import StandardScaler

# 解决中文显示问题
matplotlib.rcParams["font.family"] = ["SimHei", "Microsoft YaHei"]
matplotlib.rcParams["axes.unicode_minus"] = False


# 配置全局字体样式和按钮颜色
def setup_styles():
    style = ttk.Style()
    # 基础样式 - 统一字体颜色为黑色，文本居中
    style.configure(".",
                    font=("SimHei", 10),
                    background="#f5f5f5",
                    foreground="black",
                    justify=tk.CENTER)

    # 按钮样式 - 彩色背景+黑色文字，文本居中
    style.configure("select.TButton",
                    font=("SimHei", 10, "bold"),
                    background="#4CAF50",
                    padding=5,
                    foreground="black",
                    justify=tk.CENTER)
    style.map("select.TButton",
              foreground=[("active", "black"), ("!active", "black")],
              background=[("active", "#45a049"), ("!active", "#4CAF50")])

    style.configure("config.TButton",
                    font=("SimHei", 10, "bold"),
                    background="#2196F3",
                    padding=5,
                    foreground="black",
                    justify=tk.CENTER)
    style.map("config.TButton",
              foreground=[("active", "black"), ("!active", "black")],
              background=[("active", "#0b7dda"), ("!active", "#2196F3")])

    style.configure("calc.TButton",
                    font=("SimHei", 10, "bold"),
                    background="#FF9800",
                    padding=5,
                    foreground="black",
                    justify=tk.CENTER)
    style.map("calc.TButton",
              foreground=[("active", "black"), ("!active", "black")],
              background=[("active", "#e68a00"), ("!active", "#FF9800")])

    style.configure("save.TButton",
                    font=("SimHei", 10, "bold"),
                    background="#f44336",
                    padding=5,
                    foreground="black",
                    justify=tk.CENTER)
    style.map("save.TButton",
              foreground=[("active", "black"), ("!active", "black")],
              background=[("active", "#d32f2f"), ("!active", "#f44336")])

    # 标签样式 - 黑色字体，文本居中
    style.configure("TLabel",
                    font=("SimHei", 10),
                    background="#f5f5f5",
                    foreground="black",
                    padding=2,
                    justify=tk.CENTER,
                    anchor=tk.CENTER)

    # 标题标签样式 - 黑色字体，文本居中
    style.configure("Header.TLabel",
                    font=("SimHei", 12, "bold"),
                    background="#e0e0e0",
                    foreground="black",
                    padding=5,
                    justify=tk.CENTER,
                    anchor=tk.CENTER)

    # 树状图样式 - 黑色字体，数据居中
    style.configure("Treeview",
                    font=("SimHei", 9),
                    rowheight=22,
                    background="white",
                    fieldbackground="white",
                    foreground="black",
                    justify=tk.CENTER)
    style.configure("Treeview.Heading",
                    font=("SimHei", 10, "bold"),
                    background="#e0e0e0",
                    foreground="black",
                    justify=tk.CENTER,
                    anchor=tk.CENTER)
    style.map("Treeview",
              background=[("selected", "#a6a6a6")],
              foreground=[("selected", "white")])

    # 标签框架样式 - 黑色字体，文本居中
    style.configure("TLabelframe",
                    background="#f5f5f5",
                    padding=5,
                    justify=tk.CENTER)
    style.configure("TLabelframe.Label",
                    font=("SimHei", 10, "bold"),
                    background="#f5f5f5",
                    foreground="black",
                    padding=3,
                    justify=tk.CENTER,
                    anchor=tk.CENTER)

    # 配置 ttk.Frame 背景色
    style.configure("TFrame", background="#f5f5f5")

    return ("SimHei", 10)


class MultiAnalysisTool:
    def __init__(self, root):
        self.root = root
        self.root.title("多方法综合评价工具")
        self.root.geometry("1100x800")
        self.root.state("zoomed")
        self.root.resizable(True, True)

        # 设置主窗口背景色
        self.root.configure(bg="#f5f5f5")

        # 创建主容器（用于居中所有内容）
        self.main_container = ttk.Frame(root, padding=20)
        self.main_container.pack(expand=True, fill=tk.BOTH)

        # 创建滚动容器（嵌套在主容器中）
        # 使用 tk.Canvas 背景色来匹配整体风格
        self.main_canvas = tk.Canvas(self.main_container, bg="#f5f5f5", highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.main_container, orient="vertical", command=self.main_canvas.yview)

        # ttk.Frame 不支持 bg，通过 style 统一配置背景色
        self.scrollable_frame = ttk.Frame(self.main_canvas, padding=10, width=1000)  # 固定最大宽度，便于居中

        # 绑定滚动事件
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
        )

        self.main_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.main_canvas.configure(yscrollcommand=self.scrollbar.set)

        # 绑定鼠标滚轮事件
        self.main_canvas.bind_all("<MouseWheel>", self._on_mouse_wheel)
        self.main_canvas.bind_all("<Button-4>", lambda e: self._on_mouse_wheel(e, delta=120))
        self.main_canvas.bind_all("<Button-5>", lambda e: self._on_mouse_wheel(e, delta=-120))

        # 布局滚动容器（居中显示）
        self.main_canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # 初始化字体样式
        self.listbox_font = setup_styles()

        # 数据存储变量
        self.df = None
        self.all_columns = []
        self.index_columns = []
        self.negative_indicators = []
        self.calc_method = tk.StringVar(value="entropy")
        self.results = {}
        self.original_filename = ""
        self.current_figure = None
        self.loadings_tree = None
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

        # 初始化界面
        self._create_widgets()

    def _on_mouse_wheel(self, event, delta=None):
        """处理鼠标滚轮事件"""
        if delta is None:
            delta = event.delta
        self.main_canvas.yview_scroll(-int(delta / 120), "units")

    def _create_widgets(self):
        # 标题区域（居中显示）
        title_font = font.Font(family="SimHei", size=16, weight="bold")
        title_frame = tk.Frame(self.scrollable_frame, bg=self.colors["primary"], height=60)
        title_frame.pack(fill=tk.X, pady=(0, 20))
        title_frame.pack_propagate(False)

        title_label = tk.Label(title_frame,
                               text="多方法综合评价工具",
                               font=title_font,
                               bg=self.colors["primary"],
                               fg="white")
        title_label.pack(pady=15)

        # 顶部文件操作区（居中显示，固定宽度）
        top_frame = ttk.Frame(self.scrollable_frame, padding=10, relief=tk.RAISED, borderwidth=1)
        top_frame.pack(anchor=tk.CENTER, fill=tk.X, pady=(0, 15))

        # 文件操作区内部居中布局
        file_inner_frame = ttk.Frame(top_frame)
        file_inner_frame.pack(anchor=tk.CENTER)

        ttk.Label(file_inner_frame, text="数据文件:", font=("SimHei", 10)).pack(side=tk.LEFT, padx=8)
        self.file_path_var = tk.StringVar(value="未选择文件")
        path_label = ttk.Label(file_inner_frame, textvariable=self.file_path_var, width=50,
                               background="white", relief=tk.SUNKEN, anchor=tk.CENTER)
        path_label.pack(side=tk.LEFT, padx=8, fill=tk.X, expand=True)
        path_label.config(foreground="black")

        # 按钮框架（居中排列）
        btn_frame = ttk.Frame(file_inner_frame)
        btn_frame.pack(side=tk.LEFT, padx=8)

        ttk.Button(btn_frame, text="选择文件", command=self._load_file, style="select.TButton").pack(side=tk.LEFT,
                                                                                                     padx=4)
        ttk.Button(btn_frame, text="开始配置", command=self._start_config, style="config.TButton").pack(side=tk.LEFT,
                                                                                                        padx=4)
        ttk.Button(btn_frame, text="开始计算", command=self._calculate, style="calc.TButton").pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text="保存结果", command=self._save_results, style="save.TButton").pack(side=tk.LEFT,
                                                                                                      padx=4)

        # 计算方法选择区（居中显示）
        method_frame = ttk.LabelFrame(self.scrollable_frame, text="计算方法选择", padding=10)
        method_frame.pack(anchor=tk.CENTER, fill=tk.X, pady=(0, 15))

        # 单选按钮居中布局
        method_inner = ttk.Frame(method_frame)
        method_inner.pack(anchor=tk.CENTER)

        ttk.Label(method_inner, text="请选择评价方法:", font=("SimHei", 10)).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(
            method_inner, text="熵值法", variable=self.calc_method, value="entropy"
        ).pack(side=tk.LEFT, padx=15)
        ttk.Radiobutton(
            method_inner, text="主成分分析法", variable=self.calc_method, value="pca"
        ).pack(side=tk.LEFT, padx=15)
        ttk.Radiobutton(
            method_inner, text="基于熵权法的TOPSIS", variable=self.calc_method, value="topsis"
        ).pack(side=tk.LEFT, padx=15)

        # 配置信息与可视化区域（居中显示，固定宽高比例）
        mid_frame = ttk.Frame(self.scrollable_frame)
        mid_frame.pack(anchor=tk.CENTER, fill=tk.BOTH, expand=True, pady=(0, 15))

        # 左右分栏（保持比例，居中显示）
        mid_paned = ttk.PanedWindow(mid_frame, orient=tk.HORIZONTAL, height=400)
        mid_paned.pack(anchor=tk.CENTER, fill=tk.BOTH, expand=True)

        # 左侧：当前配置的所有指标信息展示区
        config_frame = ttk.LabelFrame(mid_paned, text="当前配置指标信息", padding=8)
        mid_paned.add(config_frame, weight=1)

        # 配置表格居中容器 - 固定高度防止尺寸错乱
        config_table_container = ttk.Frame(config_frame, height=350)
        config_table_container.pack(expand=True, fill=tk.BOTH)
        config_table_container.pack_propagate(False)

        self.config_tree = ttk.Treeview(config_table_container, show="headings", height=10)
        self.config_tree.pack(expand=True, fill=tk.BOTH, padx=5, pady=5)

        # 右侧：分析结果展示区 - 固定高度防止尺寸错乱
        self.result_visual_frame = ttk.LabelFrame(mid_paned, text="分析结果可视化", padding=8, height=350)
        mid_paned.add(self.result_visual_frame, weight=2)
        self.result_visual_frame.pack_propagate(False)

        # 记录结果展示区（居中显示，占满可用宽度）
        record_frame = ttk.LabelFrame(self.scrollable_frame, text="记录结果展示", padding=10)
        record_frame.pack(anchor=tk.CENTER, fill=tk.BOTH, expand=True, pady=(0, 15))

        # 表格容器（居中显示）- 固定高度防止尺寸错乱
        table_container = ttk.Frame(record_frame, height=300)
        table_container.pack(expand=True, fill=tk.BOTH)
        table_container.pack_propagate(False)

        # 记录表格
        self.record_tree = ttk.Treeview(table_container, show="headings", height=15)
        self.record_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 垂直滚动条
        scrollbar_y = ttk.Scrollbar(table_container, orient="vertical", command=self.record_tree.yview)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 5), pady=5)
        self.record_tree.configure(yscrollcommand=scrollbar_y.set)

        # 水平滚动条
        scrollbar_x = ttk.Scrollbar(record_frame, orient="horizontal", command=self.record_tree.xview)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=(0, 5))
        self.record_tree.configure(xscrollcommand=scrollbar_x.set)

        # 状态提示区（居中显示）
        status_frame = ttk.Frame(self.scrollable_frame, padding=8, relief=tk.FLAT, borderwidth=1)
        status_frame.pack(anchor=tk.CENTER, fill=tk.X, pady=(0, 10))

        status_inner = ttk.Frame(status_frame)
        status_inner.pack(anchor=tk.CENTER)

        self.status_var = tk.StringVar(value="请先选择数据文件")
        ttk.Label(status_inner, textvariable=self.status_var, font=("SimHei", 10, "bold")).pack(anchor=tk.CENTER)

    def _clear_visual_area(self):
        """清空分析结果展示区"""
        for widget in self.result_visual_frame.winfo_children():
            widget.destroy()
        self.current_figure = None
        self.loadings_tree = None

    def _center_window(self, window, width=None, height=None):
        """使窗口居中显示"""
        window.update_idletasks()
        if not width:
            width = window.winfo_width()
        if not height:
            height = window.winfo_height()

        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()

        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        window.geometry(f"{width}x{height}+{x}+{y}")

    def _load_file(self):
        """加载Excel数据文件"""
        from tkinter import filedialog
        file_path = filedialog.askopenfilename(
            title="选择数据文件",
            filetypes=[("Excel文件", "*.xlsx;*.xls")]
        )
        if not file_path:
            return

        try:
            self.df = pd.read_excel(file_path)
            self.all_columns = list(self.df.columns)
            self.original_filename = os.path.splitext(os.path.basename(file_path))[0]
            self.file_path_var.set(os.path.basename(file_path))
            self.status_var.set(f"数据加载成功：{len(self.df)}条记录，{len(self.all_columns)}列")
            messagebox.showinfo("成功", "数据加载完成，请点击'开始配置'进行参数设置")
        except Exception as e:
            messagebox.showerror("错误", f"文件读取失败：{str(e)}")
            self.df = None

    def _start_config(self):
        """根据选择的方法进行配置（是否需要负指标）"""
        if self.df is None:
            messagebox.showwarning("提示", "请先选择数据文件")
            return

        # 重置选择状态
        self.index_columns = []
        self.negative_indicators = []
        method = self.calc_method.get()

        # 1. 第一步：选择指标列（所有方法都需要）
        def select_index_columns():
            index_win = tk.Toplevel(self.root)
            index_win.title(f"步骤1：选择指标列（{self._get_method_name()}）")
            self._center_window(index_win, 600, 470)
            index_win.resizable(True, True)
            index_win.configure(bg="#f5f5f5")

            # 窗口内组件居中布局
            main_frame = ttk.Frame(index_win, padding=15)
            main_frame.pack(expand=True, fill=tk.BOTH)

            # 明确设置标签字体颜色为黑色
            label = ttk.Label(main_frame, text="请选择用于计算的指标列（至少1个）:",
                              font=("SimHei", 11, "bold"))
            label.pack(anchor=tk.CENTER, pady=10)
            label.config(foreground="black")  # 强制黑色字体

            frame = ttk.Frame(main_frame, relief=tk.SUNKEN, borderwidth=1)
            frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=5)

            # Listbox明确设置黑色字体
            listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE, height=15, font=self.listbox_font,
                                 bg="white", fg="black", justify=tk.CENTER)
            scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=listbox.yview)
            listbox.config(yscrollcommand=scroll.set)
            listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scroll.pack(side=tk.RIGHT, fill=tk.Y)

            for col in self.all_columns:
                listbox.insert(tk.END, col)

            # 确认选择
            def confirm_index():
                selected = [listbox.get(i) for i in listbox.curselection()]
                if not selected:
                    messagebox.showerror("错误", "请至少选择1个指标列")
                    return
                self.index_columns = selected
                index_win.destroy()

                if method == "pca":
                    messagebox.showinfo("完成", f"配置完成！\n指标列：{', '.join(selected)}")
                    self.status_var.set(f"配置完成（{self._get_method_name()}）：{len(self.index_columns)}个指标列")
                    self._show_config_info()
                else:
                    messagebox.showinfo("完成", f"已选择指标列：{', '.join(selected)}\n即将进入步骤2：选择负指标")
                    select_negative_indicators()

            btn_frame = ttk.Frame(main_frame)
            btn_frame.pack(fill=tk.X, padx=10, pady=10)

            ttk.Button(btn_frame, text="全选", command=lambda: listbox.selection_set(0, tk.END), style="config.TButton").pack(side=tk.LEFT,
                                                                                                      padx=5)
            ttk.Button(btn_frame, text="取消全选", command=lambda: listbox.selection_clear(0, tk.END), style="config.TButton").pack(
                side=tk.LEFT, padx=5)
            ttk.Button(btn_frame, text="确认指标列", command=confirm_index, style="config.TButton").pack(side=tk.RIGHT,
                                                                                                         padx=5)

        # 2. 第二步：选择负指标（熵值法和TOPSIS需要）
        def select_negative_indicators():
            neg_win = tk.Toplevel(self.root)
            neg_win.title(f"步骤2：选择负指标（{self._get_method_name()}）")
            self._center_window(neg_win, 600, 470)
            neg_win.resizable(True, True)
            neg_win.configure(bg="#f5f5f5")

            # 窗口内组件居中布局
            main_frame = ttk.Frame(neg_win, padding=15)
            main_frame.pack(expand=True, fill=tk.BOTH)

            # 明确设置标签字体颜色为黑色
            label = ttk.Label(main_frame, text="请选择负指标（其余将视为正指标）:",
                              font=("SimHei", 11, "bold"))
            label.pack(anchor=tk.CENTER, pady=10)
            label.config(foreground="black")  # 强制黑色字体

            frame = ttk.Frame(main_frame, relief=tk.SUNKEN, borderwidth=1)
            frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=5)

            # Listbox明确设置黑色字体
            listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE, height=15, font=self.listbox_font,
                                 bg="white", fg="black", justify=tk.CENTER)
            scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=listbox.yview)
            listbox.config(yscrollcommand=scroll.set)
            listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scroll.pack(side=tk.RIGHT, fill=tk.Y)

            for col in self.index_columns:
                listbox.insert(tk.END, col)

            # 确认选择
            def confirm_neg():
                selected = [listbox.get(i) for i in listbox.curselection()]
                self.negative_indicators = selected
                neg_win.destroy()

                config_info = (
                    f"配置完成！\n"
                    f"指标列：{', '.join(self.index_columns)}\n"
                    f"负指标：{', '.join(selected) if selected else '无'}"
                )
                messagebox.showinfo("全部配置完成", config_info)
                self.status_var.set(f"配置完成：{len(self.index_columns)}个指标列")
                self._show_config_info()

            btn_frame = ttk.Frame(main_frame)
            btn_frame.pack(fill=tk.X, padx=10, pady=10)

            ttk.Button(btn_frame, text="全选", command=lambda: listbox.selection_set(0, tk.END), style="config.TButton").pack(side=tk.LEFT,
                                                                                                      padx=5)
            ttk.Button(btn_frame, text="取消全选", command=lambda: listbox.selection_clear(0, tk.END), style="config.TButton").pack(
                side=tk.LEFT, padx=5)
            ttk.Button(btn_frame, text="确认负指标", command=confirm_neg, style="config.TButton").pack(side=tk.RIGHT,
                                                                                                       padx=5)

        # 开始配置流程
        select_index_columns()

    def _show_config_info(self):
        """显示当前配置的指标信息"""
        # 清空表格
        for item in self.config_tree.get_children():
            self.config_tree.delete(item)
        self.config_tree["columns"] = ()

        method = self.calc_method.get()
        if method == "pca":
            # 主成分分析法的配置信息
            columns = ("指标名称", "指标类型")
            self.config_tree["columns"] = columns
            for col in columns:
                self.config_tree.heading(col, text=col)
                self.config_tree.column(col, width=180, anchor="center")

            # 填充数据
            for col in self.index_columns:
                self.config_tree.insert("", tk.END, values=(col, "分析指标"))
        else:
            # 熵值法和TOPSIS的配置信息
            columns = ("指标名称", "指标类型")
            self.config_tree["columns"] = columns
            for col in columns:
                self.config_tree.heading(col, text=col)
                self.config_tree.column(col, width=180, anchor="center")

            # 填充数据
            for col in self.index_columns:
                indicator_type = "负指标" if col in self.negative_indicators else "正指标"
                self.config_tree.insert("", tk.END, values=(col, indicator_type))

    def _get_method_name(self):
        """获取方法中文名称"""
        method_map = {
            "entropy": "熵值法",
            "pca": "主成分分析法",
            "topsis": "基于熵权法的TOPSIS"
        }
        return method_map.get(self.calc_method.get(), "未知方法")

    def _show_weight_radar_chart(self, indicators, weights):
        """在主界面显示指标权重雷达图（居中）"""
        self._clear_visual_area()

        # 标题居中 - 明确设置黑色字体
        title_label = ttk.Label(
            self.result_visual_frame,
            text=f"{self._get_method_name()}指标权重雷达图",
            font=("SimHei", 11, "bold")
        )
        title_label.pack(anchor=tk.CENTER, pady=8)
        title_label.config(foreground="black")  # 强制黑色字体

        # 关键修改1：去掉容器固定宽高，允许自适应图表尺寸
        chart_container = ttk.Frame(self.result_visual_frame)  # 移除 width=550, height=280
        chart_container.pack(expand=True, fill=tk.BOTH, padx=10, pady=5)
        # 关键修改2：允许容器根据图表大小自适应（去掉 pack_propagate(False)）

        # 关键修改3：通过 figsize 控制图表尺寸（现在会生效）
        # figsize 单位是英寸，dpi 是每英寸像素数，最终像素尺寸=figsize*dpi
        fig = plt.figure(figsize=(3.5, 3), dpi=90)  # 改这里的 figsize 即可调整大小
        ax = fig.add_subplot(111, polar=True)
        ax.set_facecolor('#f9f9f9')

        # 绘制雷达图（原有逻辑不变）
        labels = indicators
        stats = weights
        angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
        stats = np.concatenate((stats, [stats[0]]))
        angles = angles + [angles[0]]
        labels = labels + [labels[0]]

        ax.plot(angles, stats, 'o-', linewidth=2, markersize=6, color='#2196F3')
        ax.fill(angles, stats, alpha=0.2, color='#2196F3')
        ax.set_thetagrids(np.degrees(angles), labels, fontsize=9)
        ax.set_title('指标权重分布', fontsize=10, pad=20)
        ax.set_ylim(0, max(stats) * 1.2)
        ax.grid(True, linestyle='--', alpha=0.7)

        # 关键修改4：用 subplots_adjust 调整边距（替代 tight_layout，不覆盖尺寸）
        plt.subplots_adjust(
            left=-0.28,  # 左边距（0-1之间，越小图表越靠左）
            right=0.88,  # 右边距
            top=0.85,  # 上边距（越小图表越靠上）
            bottom=0.4  # 下边距
        )

        # 关键修改5：去掉画布固定宽高，让画布自适应图表尺寸
        canvas = FigureCanvasTkAgg(fig, chart_container)
        canvas_widget = canvas.get_tk_widget()
        # 让画布填充容器，且居中显示
        canvas_widget.pack(expand=True, fill=tk.BOTH, padx=5, pady=5, anchor=tk.CENTER)
        # 移除：canvas_widget.config(width=540, height=270)  # 去掉固定尺寸锁定

        canvas.draw()
        self.current_figure = fig

    def _generate_component_names(self, components_df):
        """自动生成主成分名称"""
        component_names = []
        indicators = components_df.columns.tolist()

        for comp_idx, (comp_label, row) in enumerate(components_df.iterrows()):
            abs_loadings = row.abs()
            top2_indicators = abs_loadings.sort_values(ascending=False).index[:2].tolist()

            keywords = []
            for indicator in top2_indicators:
                stop_words = ["人均", "占比", "密度", "率", "数量", "规模", "水平"]
                words = [word for word in indicator.split() if word not in stop_words]
                if not words:
                    words = [indicator[:3]]
                keywords.append(words[0])

            unique_keywords = list(dict.fromkeys(keywords))
            name = f"主成分{comp_idx + 1}：{'+'.join(unique_keywords)}"
            component_names.append(name)

        return component_names

    def _show_pca_loadings(self, components_df):
        """在主界面显示主成分载荷矩阵（居中）"""
        self._clear_visual_area()

        # 标题居中 - 明确设置黑色字体
        title_label = ttk.Label(
            self.result_visual_frame,
            text="主成分载荷矩阵",
            font=("SimHei", 11, "bold")
        )
        title_label.pack(anchor=tk.CENTER, pady=8)
        title_label.config(foreground="black")  # 强制黑色字体

        # 表格容器（固定高度防止错乱）
        table_container = ttk.Frame(self.result_visual_frame, height=280)
        table_container.pack(expand=True, fill=tk.BOTH, padx=10, pady=5)
        table_container.pack_propagate(False)

        # 垂直滚动条
        scrollbar_y = ttk.Scrollbar(table_container, orient="vertical")
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)

        # 水平滚动条
        scrollbar_x = ttk.Scrollbar(table_container, orient="horizontal")
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

        # 创建表格
        tree = ttk.Treeview(
            table_container,
            show="headings",
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set,
            height=10  # 固定行数防止尺寸变化
        )
        tree.pack(expand=True, fill=tk.BOTH, padx=5, pady=5)

        scrollbar_y.config(command=tree.yview)
        scrollbar_x.config(command=tree.xview)

        # 生成带含义的主成分名称
        component_names = self._generate_component_names(components_df)

        # 处理列名
        columns = ["主成分（含义）"] + [str(col) for col in components_df.columns]
        tree["columns"] = tuple(columns)

        # 配置列（居中对齐）
        for col in columns:
            tree.heading(col, text=col)
            width = 180 if col == "主成分（含义）" else 120
            tree.column(col, width=width, anchor="center", stretch=True)

        # 填充数据
        for idx, (_, row) in enumerate(components_df.iterrows()):
            values = [component_names[idx]] + [round(val, 4) for val in row.values]
            tree.insert("", tk.END, values=values)

        self.loadings_tree = tree

    def _calculate(self):
        """根据选择的方法执行计算"""
        if self.df is None:
            messagebox.showwarning("提示", "请先选择数据文件")
            return
        if not self.index_columns:
            messagebox.showwarning("提示", "请先完成参数配置")
            return

        method = self.calc_method.get()
        self.results = {}

        try:
            # 显示进度窗口（居中）
            progress_win = tk.Toplevel(self.root)
            progress_win.title("计算中")
            self._center_window(progress_win, 350, 150)
            progress_win.resizable(False, False)
            progress_win.grab_set()
            progress_win.configure(bg="#f5f5f5")

            # 进度窗口内组件居中
            progress_frame = ttk.Frame(progress_win, padding=20)
            progress_frame.pack(expand=True, fill=tk.BOTH)

            # 明确设置标签字体颜色为黑色
            label = ttk.Label(progress_frame, text=f"正在使用{self._get_method_name()}计算，请稍候...",
                              font=("SimHei", 10))
            label.pack(anchor=tk.CENTER, pady=10)
            label.config(foreground="black")  # 强制黑色字体

            progress_bar = ttk.Progressbar(progress_frame, length=280, mode="determinate")
            progress_bar.pack(anchor=tk.CENTER, pady=10)
            progress_bar['value'] = 10
            progress_win.update()

            # 数据预处理
            df_clean = self.df[self.index_columns].copy()

            # 处理缺失值
            for col in self.index_columns:
                if df_clean[col].isnull().any():
                    df_clean[col].fillna(df_clean[col].mean(), inplace=True)
                    self.status_var.set(f"警告：指标'{col}'存在缺失值，已用均值填充")

            progress_bar['value'] = 30
            progress_win.update()

            # 根据方法选择计算逻辑
            if method == "entropy":
                self._calculate_entropy(df_clean, progress_bar, progress_win)
            elif method == "pca":
                self._calculate_pca(df_clean, progress_bar, progress_win)
            elif method == "topsis":
                self._calculate_topsis(df_clean, progress_bar, progress_win)

            progress_bar['value'] = 90
            progress_win.update()
            progress_win.destroy()

            # 计算完成后显示结果（更新布局防止错乱）
            self.root.update_idletasks()
            if method in ["entropy", "topsis"]:
                self._show_weight_radar_chart(
                    self.index_columns,
                    self.results['indicator']['权重(wi)'].values
                )
            elif method == "pca":
                self._show_pca_loadings(self.results['components'])

            messagebox.showinfo("成功", f"{self._get_method_name()}计算完成！")
            self._show_results()

        except Exception as e:
            progress_win.destroy()
            messagebox.showerror("错误", f"计算失败：{str(e)}")

    def _calculate_entropy(self, df_clean, progress_bar, progress_win):
        """熵值法计算"""
        # 数据标准化
        raw_data = df_clean.copy()
        norm_df = raw_data.copy()

        for col in self.index_columns:
            col_data = raw_data[col].values
            max_val = col_data.max()
            min_val = col_data.min()

            if max_val == min_val:
                norm_df[col] = 0.5
            else:
                if col in self.negative_indicators:
                    norm_df[col] = (max_val - col_data) / (max_val - min_val)
                else:
                    norm_df[col] = (col_data - min_val) / (max_val - min_val)

        progress_bar['value'] = 50
        progress_win.update()

        # 计算指标比重
        def calc_proportion(norm_df):
            prop_df = norm_df.copy()
            for col in prop_df.columns:
                total = prop_df[col].sum()
                if total == 0:
                    prop_df[col] = 1 / len(prop_df)
                else:
                    prop_df[col] = prop_df[col].apply(lambda x: x / total if x != 0 else 1e-6)
            return prop_df

        prop_df = calc_proportion(norm_df)

        # 计算熵值、差异系数和权重
        n = len(df_clean)
        k = 1 / np.log(n) if n > 1 else 1
        entropy = []
        for col in self.index_columns:
            p_values = [p if p > 0 else 1e-6 for p in prop_df[col].values]
            hi = -k * sum(p * np.log(p) for p in p_values)
            entropy.append(hi)

        g = [1 - h for h in entropy]
        sum_g = sum(g)
        weights = [gi / sum_g for gi in g] if sum_g != 0 else [1 / len(g)] * len(g)

        # 保存结果
        indicator_types = ["负指标" if col in self.negative_indicators else "正指标"
                           for col in self.index_columns]
        self.results['indicator'] = pd.DataFrame({
            '指标名称': self.index_columns,
            '熵值(Hi)': entropy,
            '差异系数(gi)': g,
            '权重(wi)': weights,
            '指标类型': indicator_types
        })
        self.results['scores'] = np.dot(norm_df.values, weights)
        self.results['norm_data'] = norm_df
        progress_bar['value'] = 70

    def _calculate_pca(self, df_clean, progress_bar, progress_win):
        """主成分分析法计算"""
        # 数据标准化
        scaler = StandardScaler()
        scaled_data = scaler.fit_transform(df_clean)
        scaled_df = pd.DataFrame(scaled_data, columns=self.index_columns)

        progress_bar['value'] = 50
        progress_win.update()

        # 执行PCA
        pca = PCA()
        pca.fit(scaled_data)

        # 计算方差贡献率
        explained_variance = pca.explained_variance_ratio_
        cumulative_variance = np.cumsum(explained_variance)

        # 确定保留主成分数量
        n_components = np.argmax(cumulative_variance >= 0.85) + 1
        if n_components < 1:
            n_components = 1

        # 重新拟合
        pca = PCA(n_components=n_components)
        principal_components = pca.fit_transform(scaled_data)

        # 综合得分
        weights = explained_variance[:n_components] / np.sum(explained_variance[:n_components])
        composite_scores = np.dot(principal_components, weights)

        # 保存主成分载荷矩阵
        components_df = pd.DataFrame(
            pca.components_,
            columns=self.index_columns,
            index=[f'主成分{i + 1}' for i in range(n_components)]
        )

        # 保存结果
        self.results['indicator'] = pd.DataFrame({
            '主成分': [f'主成分{i + 1}' for i in range(n_components)],
            '方差贡献率': explained_variance[:n_components],
            '累计贡献率': cumulative_variance[:n_components]
        })
        self.results['components'] = components_df
        self.results['scores'] = composite_scores
        self.results['scaled_data'] = scaled_df
        progress_bar['value'] = 70

    def _calculate_topsis(self, df_clean, progress_bar, progress_win):
        """基于熵权法的TOPSIS法计算"""
        # 1. 数据标准化
        raw_data = df_clean.copy()
        norm_df = raw_data.copy()

        for col in self.index_columns:
            col_data = raw_data[col].values
            max_val = col_data.max()
            min_val = col_data.min()

            if max_val == min_val:
                norm_df[col] = 0.5
            else:
                if col in self.negative_indicators:
                    norm_df[col] = (max_val - col_data) / (max_val - min_val)
                else:
                    norm_df[col] = (col_data - min_val) / (max_val - min_val)

        progress_bar['value'] = 40
        progress_win.update()

        # 2. 熵值法求权重
        def calc_entropy_weights(norm_df):
            n = len(norm_df)
            k = 1 / np.log(n) if n > 1 else 1
            entropy = []
            for col in norm_df.columns:
                p_values = norm_df[col] / norm_df[col].sum()
                p_values = [p if p > 0 else 1e-6 for p in p_values]
                hi = -k * sum(p * np.log(p) for p in p_values)
                entropy.append(hi)
            g = [1 - h for h in entropy]
            sum_g = sum(g)
            return [gi / sum_g for gi in g] if sum_g != 0 else [1 / len(g)] * len(g)

        weights = calc_entropy_weights(norm_df)

        # 3. 加权标准化矩阵
        weighted_norm = norm_df.copy()
        for i, col in enumerate(self.index_columns):
            weighted_norm[col] = weighted_norm[col] * weights[i]

        # 4. 理想解与距离计算
        ideal_best = weighted_norm.max()
        ideal_worst = weighted_norm.min()
        d_best = np.sqrt(((weighted_norm - ideal_best) ** 2).sum(axis=1))
        d_worst = np.sqrt(((weighted_norm - ideal_worst) ** 2).sum(axis=1))

        # 5. 贴近度
        scores = d_worst / (d_best + d_worst)

        # 保存结果
        self.results['indicator'] = pd.DataFrame({
            '指标名称': self.index_columns,
            '权重(wi)': weights,
            '指标类型': ["负指标" if col in self.negative_indicators else "正指标"
                         for col in self.index_columns]
        })
        self.results['scores'] = scores
        self.results['norm_data'] = norm_df
        self.results['distance'] = pd.DataFrame({
            '与正理想解距离': d_best,
            '与负理想解距离': d_worst
        })
        progress_bar['value'] = 70

    def _show_results(self):
        """展示计算结果（居中显示）"""
        method = self.calc_method.get()
        method_name = self._get_method_name()

        # 清空表格
        for item in self.record_tree.get_children():
            self.record_tree.delete(item)
        self.record_tree["columns"] = ()

        # 显示前20条记录结果
        if len(self.df) > 0:
            # 获取原始数据的所有列，加上综合得分
            record_columns = self.df.columns.tolist() + [f'{method_name}综合得分']
            record_columns = [str(col) for col in record_columns]

            # 设置列
            self.record_tree["columns"] = tuple(record_columns)

            # 配置列（居中对齐，限制最大宽度防止错乱）
            for col in record_columns:
                self.record_tree.heading(col, text=col)
                # 根据列名长度调整宽度，限制最大宽度
                if col == f'{method_name}综合得分':
                    width = 150
                else:
                    width = min(len(col) * 12 + 20, 120)  # 限制最大宽度120
                self.record_tree.column(col, width=width, anchor="center", stretch=True)

            # 取前20条数据
            top20_df = self.df.head(20).copy()
            score_col = f'{method_name}综合得分'
            top20_df[score_col] = self.results['scores'][:20]

            # 填充记录数据
            for _, row in top20_df.iterrows():
                values = []
                for col in record_columns:
                    if col == score_col:
                        values.append(round(row[col], 4))
                    else:
                        val = str(row[col])
                        # 超长文本截断，保证表格美观
                        if len(val) > 15:
                            val = val[:12] + "..."
                        values.append(val)
                self.record_tree.insert("", tk.END, values=values)

        # 更新布局防止错乱
        self.root.update_idletasks()

    def _save_results(self):
        """保存结果到Excel"""
        if not self.results or not self.original_filename:
            messagebox.showwarning("提示", "请先完成数据加载和计算")
            return

        method = self.calc_method.get()
        method_name = self._get_method_name()

        try:
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            new_filename = f"{current_time}_{self.original_filename}_{method_name}结果.xlsx"
            save_path = os.path.join(os.getcwd(), new_filename)

            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                # 1. 指标分析结果
                self.results['indicator'].to_excel(writer, sheet_name=f'{method_name}指标结果', index=False)

                # 2. 原始数据+综合得分
                df_with_score = self.df.copy()
                df_with_score[f'{method_name}综合得分'] = self.results['scores']
                df_with_score.to_excel(writer, sheet_name='原始数据+综合得分', index=False)

                # 3. 方法特定数据
                if method == "entropy":
                    self.results['norm_data'].to_excel(writer, sheet_name='标准化数据', index=False)
                elif method == "pca":
                    self.results['scaled_data'].to_excel(writer, sheet_name='标准化数据', index=False)
                    self.results['components'].to_excel(writer, sheet_name='主成分载荷矩阵', index=True)
                elif method == "topsis":
                    self.results['norm_data'].to_excel(writer, sheet_name='标准化数据', index=False)
                    self.results['distance'].to_excel(writer, sheet_name='TOPSIS距离结果', index=False)

            messagebox.showinfo("成功", f"{method_name}结果已保存至：\n{save_path}")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败：{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = MultiAnalysisTool(root)
    root.mainloop()