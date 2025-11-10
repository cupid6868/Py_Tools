import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
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
    style.configure(".", font=("SimHei", 10))

    # 按钮样式
    style.configure("select.TButton",
                    font=("SimHei", 10, "bold"),
                    background="#4CAF50")
    style.map("select.TButton",
              foreground=[("active", "black"), ("!active", "black")],
              background=[("active", "#45a049"), ("!active", "#4CAF50")])

    style.configure("config.TButton",
                    font=("SimHei", 10, "bold"),
                    background="#2196F3")
    style.map("config.TButton",
              foreground=[("active", "black"), ("!active", "black")],
              background=[("active", "#0b7dda"), ("!active", "#2196F3")])

    style.configure("calc.TButton",
                    font=("SimHei", 10, "bold"),
                    background="#FF9800")
    style.map("calc.TButton",
              foreground=[("active", "black"), ("!active", "black")],
              background=[("active", "#e68a00"), ("!active", "#FF9800")])

    style.configure("save.TButton",
                    font=("SimHei", 10, "bold"),
                    background="#f44336")
    style.map("save.TButton",
              foreground=[("active", "black"), ("!active", "black")],
              background=[("active", "#d32f2f"), ("!active", "#f44336")])

    style.configure("TLabel", font=("SimHei", 10))
    style.configure("Treeview", font=("SimHei", 9))
    style.configure("Treeview.Heading", font=("SimHei", 10, "bold"))
    return ("SimHei", 10)


class MultiAnalysisTool:
    def __init__(self, root):
        self.root = root
        self.root.title("多方法综合评价工具")
        self.root.geometry("1000x750")  # 主窗口固定大小
        self.root.resizable(True, True)  # 禁止窗口缩放

        # 创建主容器（非滚动，避免整体布局变动）
        self.main_frame = ttk.Frame(root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 初始化字体样式
        self.listbox_font = setup_styles()

        # 数据存储变量
        self.df = None  # 原始数据
        self.all_columns = []
        self.index_columns = []  # 指标列
        self.negative_indicators = []  # 负指标（熵值法和TOPSIS用）
        self.calc_method = tk.StringVar(value="entropy")  # 计算方法：entropy/pca/topsis
        self.results = {}  # 存储计算结果
        self.original_filename = ""  # 原始文件名
        self.current_figure = None  # 当前图表
        self.loadings_tree = None  # 载荷矩阵表格

        # 初始化界面
        self._create_widgets()

    def _create_widgets(self):
        # 顶部文件操作区
        top_frame = ttk.Frame(self.main_frame, padding=3)
        top_frame.pack(fill=tk.X)

        ttk.Label(top_frame, text="数据文件:").pack(side=tk.LEFT, padx=2)
        self.file_path_var = tk.StringVar(value="未选择文件")
        ttk.Label(top_frame, textvariable=self.file_path_var, width=35).pack(side=tk.LEFT, padx=2)

        # 带颜色的按钮
        ttk.Button(top_frame, text="选择文件", command=self._load_file, style="select.TButton").pack(side=tk.LEFT,
                                                                                                     padx=2)
        ttk.Button(top_frame, text="开始配置", command=self._start_config, style="config.TButton").pack(side=tk.LEFT,
                                                                                                        padx=2)
        ttk.Button(top_frame, text="开始计算", command=self._calculate, style="calc.TButton").pack(side=tk.LEFT, padx=2)
        ttk.Button(top_frame, text="保存结果", command=self._save_results, style="save.TButton").pack(side=tk.LEFT,
                                                                                                      padx=2)

        # 计算方法选择区
        method_frame = ttk.Frame(self.main_frame, padding=3)
        method_frame.pack(fill=tk.X)
        ttk.Label(method_frame, text="计算方法:").pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(
            method_frame, text="熵值法", variable=self.calc_method, value="entropy"
        ).pack(side=tk.LEFT, padx=3)
        ttk.Radiobutton(
            method_frame, text="主成分分析法", variable=self.calc_method, value="pca"
        ).pack(side=tk.LEFT, padx=3)
        ttk.Radiobutton(
            method_frame, text="基于熵权法的TOPSIS", variable=self.calc_method, value="topsis"
        ).pack(side=tk.LEFT, padx=3)

        # 配置信息与可视化区域 - 固定高度
        mid_frame = ttk.Frame(self.main_frame)
        mid_frame.pack(fill=tk.X, pady=2)
        mid_frame.configure(height=300)  # 固定中间区域高度
        mid_frame.pack_propagate(False)  # 禁止中间区域随内容缩放

        # 左右分栏
        mid_paned = ttk.PanedWindow(mid_frame, orient=tk.HORIZONTAL, height=290)
        mid_paned.pack(fill=tk.BOTH, expand=True, padx=3, pady=2)

        # 左侧：当前配置的所有指标信息展示区
        config_frame = ttk.LabelFrame(mid_paned, text="当前配置指标信息", padding=3, height=280)
        mid_paned.add(config_frame, weight=1)
        config_frame.pack_propagate(False)  # 禁止配置区缩放
        self.config_tree = ttk.Treeview(config_frame, show="headings", height=6)
        self.config_tree.pack(fill=tk.BOTH, expand=True)

        # 右侧：分析结果展示区（方法特定图表/表格）
        self.result_visual_frame = ttk.LabelFrame(mid_paned, text="分析结果可视化", padding=3, height=280)
        mid_paned.add(self.result_visual_frame, weight=2)
        self.result_visual_frame.pack_propagate(False)  # 禁止结果可视化区缩放

        # 前十条记录结果展示区 - 完全固定大小
        record_frame = ttk.LabelFrame(self.main_frame, text="前十条记录结果", padding=5)
        record_frame.pack(fill=tk.X, pady=4)
        record_frame.configure(height=480)  # 固定记录区总高度
        record_frame.pack_propagate(False)  # 关键：禁止记录区随内容变化

        # 表格容器（用于放置表格和滚动条）- 固定尺寸
        table_container = ttk.Frame(record_frame, width=980, height=260)  # 固定像素尺寸
        table_container.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        table_container.pack_propagate(False)  # 禁止容器随内容缩放

        # 记录表格（固定显示行数）
        self.record_tree = ttk.Treeview(table_container, show="headings", height=12)  # height=12行
        self.record_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 垂直滚动条
        scrollbar_y = ttk.Scrollbar(table_container, orient="vertical", command=self.record_tree.yview)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.record_tree.configure(yscrollcommand=scrollbar_y.set)

        # 水平滚动条
        scrollbar_x = ttk.Scrollbar(record_frame, orient="horizontal", command=self.record_tree.xview)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.record_tree.configure(xscrollcommand=scrollbar_x.set)

        # 状态提示
        self.status_var = tk.StringVar(value="请先选择数据文件")
        ttk.Label(self.main_frame, textvariable=self.status_var, foreground="blue").pack(pady=3, anchor=tk.W)

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
            self._center_window(index_win, 450, 350)

            ttk.Label(index_win, text="请选择用于计算的指标列（至少1个）:").pack(anchor=tk.W, padx=8, pady=3)

            frame = ttk.Frame(index_win)
            frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=3)

            listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE, height=10, font=self.listbox_font)
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

                # 主成分分析法不需要负指标
                if method == "pca":
                    messagebox.showinfo("完成", f"配置完成！\n指标列：{', '.join(selected)}")
                    self.status_var.set(f"配置完成（{self._get_method_name()}）：{len(self.index_columns)}个指标列")
                    self._show_config_info()
                else:
                    messagebox.showinfo("完成", f"已选择指标列：{', '.join(selected)}\n即将进入步骤2：选择负指标")
                    select_negative_indicators()

            btn_frame = ttk.Frame(index_win)
            btn_frame.pack(fill=tk.X, padx=8, pady=3)

            ttk.Button(btn_frame, text="确认指标列", command=confirm_index).pack(side=tk.RIGHT, padx=3)
            ttk.Button(btn_frame, text="全选", command=lambda: listbox.selection_set(0, tk.END)).pack(side=tk.LEFT,
                                                                                                      padx=3)
            ttk.Button(btn_frame, text="取消全选", command=lambda: listbox.selection_clear(0, tk.END)).pack(
                side=tk.LEFT, padx=3)

        # 2. 第二步：选择负指标（熵值法和TOPSIS需要）
        def select_negative_indicators():
            neg_win = tk.Toplevel(self.root)
            neg_win.title(f"步骤2：选择负指标（{self._get_method_name()}）")
            self._center_window(neg_win, 450, 350)

            ttk.Label(neg_win, text="请选择负指标（其余将视为正指标）:").pack(anchor=tk.W, padx=8, pady=3)

            frame = ttk.Frame(neg_win)
            frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=3)

            listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE, height=10, font=self.listbox_font)
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

            btn_frame = ttk.Frame(neg_win)
            btn_frame.pack(fill=tk.X, padx=8, pady=3)

            ttk.Button(btn_frame, text="确认负指标", command=confirm_neg).pack(side=tk.RIGHT, padx=3)
            ttk.Button(btn_frame, text="全选", command=lambda: listbox.selection_set(0, tk.END)).pack(side=tk.LEFT,
                                                                                                      padx=3)
            ttk.Button(btn_frame, text="取消全选", command=lambda: listbox.selection_clear(0, tk.END)).pack(
                side=tk.LEFT, padx=3)

        # 开始配置流程
        select_index_columns()

    def _show_config_info(self):
        """显示当前配置的指标信息"""
        # 清空表格
        for item in self.config_tree.get_children():
            self.config_tree.delete(item)
        self.config_tree["columns"] = ()  # 重置为空白元组

        method = self.calc_method.get()
        if method == "pca":
            # 主成分分析法的配置信息
            columns = ("指标名称", "指标类型")
            self.config_tree["columns"] = columns
            for col in columns:
                self.config_tree.heading(col, text=col)
                self.config_tree.column(col, width=120, anchor="center")
        else:
            # 熵值法和TOPSIS的配置信息
            columns = ("指标名称", "指标类型")
            self.config_tree["columns"] = columns
            for col in columns:
                self.config_tree.heading(col, text=col)
                self.config_tree.column(col, width=120, anchor="center")

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
        """在主界面显示指标权重雷达图"""
        self._clear_visual_area()

        # 设置标题
        ttk.Label(
            self.result_visual_frame,
            text=f"{self._get_method_name()}指标权重雷达图",
            font=("SimHei", 11, "bold")
        ).pack(pady=3)

        # 创建图表（适应固定区域大小）
        fig = plt.figure(figsize=(5, 4), dpi=90)
        ax = fig.add_subplot(111, polar=True)

        # 绘制雷达图
        labels = indicators
        stats = weights
        angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
        stats = np.concatenate((stats, [stats[0]]))
        angles = angles + [angles[0]]
        labels = labels + [labels[0]]

        ax.plot(angles, stats, 'o-', linewidth=1.5, markersize=5)
        ax.fill(angles, stats, alpha=0.25)
        ax.set_thetagrids(np.degrees(angles), labels, fontsize=9)
        ax.set_title('指标权重分布', fontsize=10)
        ax.set_ylim(0, max(stats) * 1.2)
        ax.grid(True, linestyle='--', alpha=0.7)

        # 紧凑布局
        plt.tight_layout()

        # 将图表嵌入主界面
        canvas = FigureCanvasTkAgg(fig, self.result_visual_frame)
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        canvas.draw()
        self.current_figure = fig

    def _generate_component_names(self, components_df):
        """自动生成主成分名称"""
        component_names = []
        indicators = components_df.columns.tolist()

        for comp_idx, (comp_label, row) in enumerate(components_df.iterrows()):
            # 获取当前主成分载荷绝对值最大的前2个指标
            abs_loadings = row.abs()
            top2_indicators = abs_loadings.sort_values(ascending=False).index[:2].tolist()

            # 提取核心关键词
            keywords = []
            for indicator in top2_indicators:
                stop_words = ["人均", "占比", "密度", "率", "数量", "规模", "水平"]
                words = [word for word in indicator.split() if word not in stop_words]
                if not words:
                    words = [indicator[:3]]
                keywords.append(words[0])

            # 去重并生成名称
            unique_keywords = list(dict.fromkeys(keywords))
            name = f"主成分{comp_idx + 1}：{'+'.join(unique_keywords)}"
            component_names.append(name)

        return component_names

    def _show_pca_loadings(self, components_df):
        """在主界面显示主成分载荷矩阵"""
        self._clear_visual_area()

        # 设置标题
        ttk.Label(
            self.result_visual_frame,
            text="主成分载荷矩阵",
            font=("SimHei", 11, "bold")
        ).pack(pady=3)

        # 创建滚动条容器
        frame = ttk.Frame(self.result_visual_frame)
        frame.pack(fill=tk.BOTH, expand=True)

        # 垂直滚动条
        scrollbar_y = ttk.Scrollbar(frame, orient="vertical")
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)

        # 水平滚动条
        scrollbar_x = ttk.Scrollbar(frame, orient="horizontal")
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

        # 创建表格
        tree = ttk.Treeview(
            frame,
            show="headings",
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set
        )
        tree.pack(fill=tk.BOTH, expand=True)

        scrollbar_y.config(command=tree.yview)
        scrollbar_x.config(command=tree.xview)

        # 生成带含义的主成分名称
        component_names = self._generate_component_names(components_df)

        # 处理列名
        columns = ["主成分（含义）"] + [str(col) for col in components_df.columns]
        tree["columns"] = tuple(columns)

        # 配置列
        for col in columns:
            tree.heading(col, text=col)
            width = 150 if col == "主成分（含义）" else 110
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
        self.results = {}  # 清空之前的结果

        try:
            # 显示进度窗口
            progress_win = tk.Toplevel(self.root)
            progress_win.title("计算中")
            self._center_window(progress_win, 280, 100)

            ttk.Label(progress_win, text=f"正在使用{self._get_method_name()}计算，请稍候...").pack(pady=8)
            progress_bar = ttk.Progressbar(progress_win, length=230, mode="determinate")
            progress_bar.pack(pady=8)
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

            # 计算完成后在主界面显示相应内容
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
        # 数据标准化（均值为0，方差为1）
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

        # 确定保留主成分数量（累计贡献率>85%）
        n_components = np.argmax(cumulative_variance >= 0.85) + 1
        if n_components < 1:
            n_components = 1

        # 重新拟合
        pca = PCA(n_components=n_components)
        principal_components = pca.fit_transform(scaled_data)

        # 综合得分（按贡献率加权）
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
        self.results['components'] = components_df  # 主成分载荷矩阵
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

        # 5. 贴近度（综合得分）
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
        """展示计算结果（仅显示前十条记录结果）"""
        method = self.calc_method.get()
        method_name = self._get_method_name()

        # 清空表格
        for item in self.record_tree.get_children():
            self.record_tree.delete(item)
        self.record_tree["columns"] = ()  # 重置列为空元组

        # 显示前十条记录结果
        if len(self.df) > 0:
            # 获取原始数据的所有列，加上综合得分
            record_columns = self.df.columns.tolist() + [f'{method_name}综合得分']
            record_columns = [str(col) for col in record_columns]

            # 直接设置列
            self.record_tree["columns"] = tuple(record_columns)

            # 配置列标题和宽度（固定列宽）
            for col in record_columns:
                self.record_tree.heading(col, text=col)
                self.record_tree.column(col, width=100, anchor="center", stretch=False)  # 禁止列拉伸

            # 取前十条数据
            top10_df = self.df.head(20).copy()
            score_col = f'{method_name}综合得分'
            top10_df[score_col] = self.results['scores'][:20]

            # 填充记录数据
            for _, row in top10_df.iterrows():
                values = [str(row[col]) if col != score_col else round(row[score_col], 4)
                          for col in record_columns[:-1]]
                values.append(round(row[score_col], 4))
                self.record_tree.insert("", tk.END, values=values)

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
