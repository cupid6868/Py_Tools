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
import threading
import time
import traceback
from concurrent.futures import ThreadPoolExecutor, as_completed
import warnings

warnings.filterwarnings('ignore')

# ===================== RAGA-PPC 全局参数配置（固定值） =====================
POPULATION_SIZE = 400  # 降低种群规模减少内存占用
CROSSOVER_RATE = 0.8
MUTATION_RATE = 0.2  # 降低变异率减少计算量
ELITE_INDIVIDUALS = 20  # 减少精英个体数
ACCELERATE_TIMES = 7  # 减少加速次数
# ===========================================================================

# 解决中文显示问题
matplotlib.rcParams["font.family"] = ["SimHei", "Microsoft YaHei"]
matplotlib.rcParams["axes.unicode_minus"] = False

# 全局线程池（降低线程数，减少CPU竞争）
GLOBAL_THREAD_POOL = ThreadPoolExecutor(
    max_workers=max(2, os.cpu_count()),  # 从2*CPU核数改为CPU核数
    thread_name_prefix="RAGA_WORKER"
)

# 全局日志缓存（保证控制台和窗口日志一致）
GLOBAL_LOG_CACHE = []
LOG_LOCK = threading.Lock()


# --- 优化的日志函数（保证控制台和窗口日志一致） ---
def log(message, log_text_widget=None):
    """日志函数，支持输出到控制台和日志窗口，增加时间戳精度，保证输出一致"""
    timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]  # 毫秒级时间戳
    log_msg = f"[{timestamp}] {message}"

    # 1. 线程安全写入全局缓存
    with LOG_LOCK:
        GLOBAL_LOG_CACHE.append(log_msg)

    # 2. 输出到控制台
    print(log_msg)

    # 3. 输出到日志窗口（线程安全）
    if log_text_widget and log_text_widget.winfo_exists():
        try:
            # 使用after保证tkinter线程安全
            def safe_insert():
                if log_text_widget.winfo_exists():
                    log_text_widget.insert(tk.END, log_msg + "\n")
                    log_text_widget.see(tk.END)
                    log_text_widget.update_idletasks()

            log_text_widget.after(0, safe_insert)
        except Exception as e:
            print(f"日志窗口输出失败: {e}")


# --- 优化的核心计算函数 ---
def core_projection_index_vectorized(a, norm_data, window_radius, n):
    """向量化优化：减少内存拷贝，优化计算逻辑"""
    a = np.asarray(a).ravel()
    # 投影方向单位化约束（提前返回减少计算）
    norm_a = np.linalg.norm(a)
    if norm_a < 1e-10:
        return 0.0
    a = a / norm_a

    # 1. 计算投影值 Z = X * a（使用矩阵乘法优化）
    z = norm_data @ a  # 替代np.dot，性能更优

    # 2. 计算标准差 Sz
    sz = np.std(z, ddof=1)
    if sz < 1e-10:
        return 0.0

    # 3. 优化局部密度 Dz 计算（减少内存占用）
    r = window_radius if window_radius else 0.1 * sz
    z_sorted = np.sort(z)

    # 滑动窗口优化：使用广播代替全矩阵计算
    dz = 0.0
    for i in range(n):
        # 只计算当前点附近的范围，减少计算量
        mask = np.abs(z_sorted - z_sorted[i]) <= r
        dz += np.sum(r - np.abs(z_sorted[mask] - z_sorted[i]))

    return sz * dz


def batch_evaluate_population(pop, norm_data, window_radius, n, stop_flag):
    """优化批量评估：分批次计算，减少内存占用"""
    if stop_flag.is_set():
        return np.zeros(len(pop))

    fitness = np.zeros(len(pop))
    batch_size = 50  # 分批次计算，避免一次性占用过多内存
    completed = 0

    # 分批次提交任务
    for batch_start in range(0, len(pop), batch_size):
        if stop_flag.is_set():
            break

        batch_end = min(batch_start + batch_size, len(pop))
        batch_pop = pop[batch_start:batch_end]

        # 提交当前批次任务
        futures = {
            GLOBAL_THREAD_POOL.submit(core_projection_index_vectorized, ind, norm_data, window_radius, n): idx
            for idx, ind in enumerate(batch_pop, start=batch_start)
        }

        for future in as_completed(futures):
            if stop_flag.is_set():
                for f in futures:
                    if not f.done():
                        f.cancel()
                break

            idx = futures[future]
            try:
                fitness[idx] = future.result()
                completed += 1
            except Exception as e:
                log(f"个体 {idx} 计算失败: {str(e)}", None)
                fitness[idx] = 0.0

    return fitness


class RAGA_PPC_Engine:
    def __init__(self, data, progress_callback=None, log_text_widget=None):
        self.log_text_widget = log_text_widget
        log(f"初始化 RAGA 模型...", self.log_text_widget)
        log(f"固定参数配置：种群规模={POPULATION_SIZE}, 交叉概率={CROSSOVER_RATE}, 变异概率={MUTATION_RATE}, 优秀个体数={ELITE_INDIVIDUALS}, 加速次数={ACCELERATE_TIMES}",
            self.log_text_widget)

        self.raw_data = np.array(data, dtype=float)
        # 增加数据维度校验
        if len(self.raw_data.shape) != 2:
            raise ValueError(f"输入数据必须是2维数组，当前形状：{self.raw_data.shape}")
        if self.raw_data.shape[1] == 0:
            raise ValueError("输入数据不能为空（无指标列）")

        self.column_names = data.columns.tolist()
        self.progress_callback = progress_callback

        # 使用线程安全的停止标志
        self.stop_flag = threading.Event()

        # Q值稳定检测相关
        self.best_q_history = []
        self.stable_iter_count = 0
        self.STABLE_THRESHOLD = 8  # 降低稳定阈值，加快收敛
        self.DECIMAL_DIGITS = 4

        # 1. 极差归一化（优化内存使用）
        self.min_val = self.raw_data.min(axis=0)
        self.max_val = self.raw_data.max(axis=0)
        diff = self.max_val - self.min_val
        diff[diff == 0] = 1.0
        self.norm_data = (self.raw_data - self.min_val) / diff

        self.n, self.m = self.norm_data.shape
        log(f"数据预处理完成。样本量: {self.n}, 维度: {self.m}", self.log_text_widget)

        # 预分配内存（优化）
        self._preallocate_buffers()

    def _preallocate_buffers(self):
        """预分配内存，减少计算过程中的内存申请"""
        self._fitness_buffer = np.zeros(POPULATION_SIZE)
        self._elite_buffer = np.zeros((ELITE_INDIVIDUALS, self.m))
        self._new_pop_buffer = np.zeros((POPULATION_SIZE, self.m))

    def _check_q_stability(self, current_q):
        """检查Q值是否稳定"""
        current_q_rounded = round(current_q, self.DECIMAL_DIGITS)

        if len(self.best_q_history) > 0:
            last_q_rounded = round(self.best_q_history[-1], self.DECIMAL_DIGITS)

            if current_q_rounded == last_q_rounded:
                self.stable_iter_count += 1
                if self.stable_iter_count >= self.STABLE_THRESHOLD:
                    log(f"Q值连续{self.STABLE_THRESHOLD}次迭代保持稳定（{current_q_rounded}），提前结束计算",
                        self.log_text_widget)
                    return True
            else:
                self.stable_iter_count = 0

        self.best_q_history.append(current_q)
        return False

    def _vectorized_selection(self, pop, fitness, n_pop):
        """向量化选择操作（优化内存）"""
        # 精英保留（向量化）
        elite_size = ELITE_INDIVIDUALS
        elite_indices = np.argpartition(fitness, -elite_size)[-elite_size:]
        self._elite_buffer[:] = pop[elite_indices]

        # 锦标赛选择（优化）
        tournament_size = 2  # 降低锦标赛规模，减少计算
        n_tournament = n_pop - elite_size

        # 向量化生成锦标赛索引
        tournament_idx = np.random.randint(0, len(pop), (n_tournament, tournament_size))
        tournament_fitness = fitness[tournament_idx]
        winner_idx = np.argmax(tournament_fitness, axis=1)

        # 向量化选择获胜者
        selected_idx = tournament_idx[np.arange(n_tournament), winner_idx]
        self._new_pop_buffer[:elite_size] = self._elite_buffer
        self._new_pop_buffer[elite_size:n_pop] = pop[selected_idx]

        return self._new_pop_buffer[:n_pop]

    def _vectorized_crossover(self, pop):
        """向量化交叉操作（优化）"""
        n_pop, m = pop.shape

        # 生成交叉掩码
        crossover_mask = np.random.rand(n_pop) < CROSSOVER_RATE
        crossover_idx = np.where(crossover_mask)[0]
        if len(crossover_idx) % 2 != 0:
            crossover_idx = crossover_idx[:-1]

        if len(crossover_idx) < 2:
            return pop

        # 配对
        parent1_idx = crossover_idx[::2]
        parent2_idx = crossover_idx[1::2]

        # 向量化生成交叉因子（优化）
        alpha = np.random.rand(len(parent1_idx), 1)

        # 原地计算，减少内存拷贝
        pop[parent1_idx] = alpha * pop[parent1_idx] + (1 - alpha) * pop[parent2_idx]
        pop[parent2_idx] = (1 - alpha) * pop[parent1_idx] + alpha * pop[parent2_idx]

        return pop

    def _vectorized_mutation(self, pop, lb=0, ub=1):
        """向量化变异操作（终极修复：兼容标量/数组上下限，无广播问题）"""
        n_pop, m = pop.shape

        # 1. 统一处理lb/ub：转为数组格式，方便后续操作
        if isinstance(lb, (int, float)):
            lb_arr = np.full(m, lb)  # 标量转数组
        else:
            lb_arr = np.asarray(lb).ravel()[:m]  # 数组标准化

        if isinstance(ub, (int, float)):
            ub_arr = np.full(m, ub)  # 标量转数组
        else:
            ub_arr = np.asarray(ub).ravel()[:m]  # 数组标准化

        # 2. 生成变异掩码
        mutation_mask = np.random.rand(n_pop, m) < MUTATION_RATE

        # 3. 按维度生成随机数（彻底避免广播）
        for dim in range(m):
            dim_mask = mutation_mask[:, dim]
            if not np.any(dim_mask):
                continue
            # 为当前维度生成匹配的随机数
            pop[dim_mask, dim] = np.random.uniform(
                lb_arr[dim], ub_arr[dim], size=np.sum(dim_mask)
            )

        return pop

    def solve(self, n_gen=30, window_radius=None):
        """优化后的RAGA求解过程（兼容所有numpy版本）"""
        log(f"启动固定参数版RAGA算法: 全局线程池大小={GLOBAL_THREAD_POOL._max_workers}", self.log_text_widget)
        log(f"固定参数：种群规模={POPULATION_SIZE}, 交叉概率={CROSSOVER_RATE}, 变异概率={MUTATION_RATE}, 优秀个体数={ELITE_INDIVIDUALS}, 加速次数={ACCELERATE_TIMES}",
            self.log_text_widget)
        log(f"Q值稳定检测: 小数点后{self.DECIMAL_DIGITS}位，连续{self.STABLE_THRESHOLD}次稳定则提前结束",
            self.log_text_widget)

        # 初始搜索范围（保留数组格式，支持维度级上下限）
        lb = np.zeros(self.m)
        ub = np.ones(self.m)

        best_a = None
        best_q = -1.0
        total_iter = n_gen * ACCELERATE_TIMES
        current_iter = 0

        for acc in range(ACCELERATE_TIMES):
            if self.stop_flag.is_set():
                log("计算被终止", self.log_text_widget)
                break

            # 修复：统一用np.min/np.max，兼容数组/标量
            lb_min = np.min(lb)
            ub_max = np.max(ub)
            log(f"\n=== 第 {acc + 1}/{ACCELERATE_TIMES} 轮加速搜索 (当前搜索区间: [{lb_min:.4f}, {ub_max:.4f}]) ===",
                self.log_text_widget)

            # 初始化种群 (优化内存)
            pop = np.random.uniform(lb, ub, (POPULATION_SIZE, self.m))

            for gen in range(n_gen):
                if self.stop_flag.is_set():
                    break

                current_iter += 1
                # 更新进度
                if self.progress_callback:
                    progress = 20 + (current_iter / total_iter) * 70
                    self.progress_callback(min(progress, 90))

                # 1. 批量线程评估适应度
                fitness = batch_evaluate_population(
                    pop, self.norm_data, window_radius, self.n, self.stop_flag
                )

                # 记录本轮最优
                idx_best = np.argmax(fitness)
                current_best_q = fitness[idx_best]
                current_avg_q = np.mean(fitness[fitness > 0])  # 计算有效平均Q值

                if current_best_q > best_q:
                    best_q = current_best_q
                    best_a = pop[idx_best].copy()

                # 增强日志：每轮都打印详细信息
                log(f"  加速轮{acc + 1} | 迭代{gen + 1}/{n_gen} | 本轮最优Q: {current_best_q:.6f} | "
                    f"全局最优Q: {best_q:.6f} | 平均Q: {current_avg_q:.6f} | "
                    f"稳定计数: {self.stable_iter_count}/{self.STABLE_THRESHOLD}",
                    self.log_text_widget)

                # 检查Q值稳定性
                if self._check_q_stability(best_q):
                    self.stop_flag.set()
                    break

                # 2. 向量化选择
                pop = self._vectorized_selection(pop, fitness, POPULATION_SIZE)

                # 3. 向量化交叉
                pop = self._vectorized_crossover(pop)

                # 4. 向量化变异
                pop = self._vectorized_mutation(pop, lb=lb, ub=ub)

            if self.stop_flag.is_set():
                break

            if best_a is not None:
                # 加速步骤：缩小搜索区间（优化）
                delta = (ub - lb) * 0.3
                lb = np.maximum(0, best_a - delta)
                ub = np.minimum(1, best_a + delta)
                # 修复：日志输出用np.min/np.max
                log(f"  第 {acc + 1} 轮加速完成，新搜索区间: [{np.min(lb):.4f}, {np.max(ub):.4f}]",
                    self.log_text_widget)

            # 释放内存（兼容所有numpy版本的写法）
            del pop
            del fitness
            # 移除 np.clear_cache()，替换为以下兼容写法
            import gc
            gc.collect()  # 强制Python垃圾回收，释放内存

        if best_a is None:
            best_a = np.ones(self.m) / self.m

        # 归一化投影向量
        best_a = best_a / np.linalg.norm(best_a)
        scores = np.dot(self.norm_data, best_a)

        # ========== 核心修改1：RAGA-PPC得分归一化到0-1区间 ==========
        scores_min = np.min(scores)
        scores_max = np.max(scores)
        # 防止除零错误
        if scores_max - scores_min < 1e-10:
            normalized_scores = np.zeros_like(scores)
        else:
            normalized_scores = (scores - scores_min) / (scores_max - scores_min)

        log(f"\n===== 计算完成 ======", self.log_text_widget)
        log(f"最终最优投影指标 Q: {best_q:.6f}", self.log_text_widget)
        log(f"最优投影方向向量: {np.round(best_a, 4)}", self.log_text_widget)
        log(f"原始投影得分范围: [{scores_min:.6f}, {scores_max:.6f}]", self.log_text_widget)
        log(f"归一化后得分范围: [{np.min(normalized_scores):.6f}, {np.max(normalized_scores):.6f}]",
            self.log_text_widget)

        return best_a, normalized_scores  # 返回归一化后的得分

    def stop_calculation(self):
        """安全停止计算"""
        self.stop_flag.set()
        log("计算停止指令已执行", self.log_text_widget)


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
        self.df = None  # 原始数据完整存储
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

        # RAGA-PPC 新增变量
        self.raga_engine = None
        self.raga_thread = None
        self.is_calculating = False
        self.log_window = None
        self.log_text_widget = None
        self.stop_flag = threading.Event()
        self.raga_n_gen = tk.IntVar(value=10)  # 减少默认迭代次数
        self.raga_window_radius = tk.DoubleVar(value=0.0)

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
        self.calc_button = ttk.Button(btn_frame, text="开始计算", command=self._calculate, style="calc.TButton")
        self.calc_button.pack(side=tk.LEFT, padx=4)
        self.stop_button = ttk.Button(btn_frame, text="终止计算", command=self._stop_calculate, style="save.TButton")
        self.stop_button.pack(side=tk.LEFT, padx=4)
        self.stop_button.config(state=tk.DISABLED)
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
        # 新增 RAGA-PPC 单选按钮
        ttk.Radiobutton(
            method_inner, text="RAGA-PPC投影寻踪模型", variable=self.calc_method, value="raga_ppc"
        ).pack(side=tk.LEFT, padx=15)

        # 新增 RAGA-PPC 参数说明区域
        raga_info_frame = ttk.LabelFrame(self.scrollable_frame, text="RAGA-PPC固定参数说明", padding=10)
        raga_info_frame.pack(anchor=tk.CENTER, fill=tk.X, pady=(0, 15))

        info_inner = ttk.Frame(raga_info_frame)
        info_inner.pack(anchor=tk.CENTER)

        param_info = (
            f"种群规模: {POPULATION_SIZE} | 交叉概率: {CROSSOVER_RATE} | 变异概率: {MUTATION_RATE} | "
            f"优秀个体数: {ELITE_INDIVIDUALS} | 加速次数: {ACCELERATE_TIMES}"
        )
        ttk.Label(info_inner, text=param_info, font=("SimHei", 10, "bold"), foreground="#e5383b").pack(anchor=tk.CENTER,
                                                                                                       pady=5)

        # RAGA-PPC 可配置参数
        iter_frame = ttk.Frame(raga_info_frame)
        iter_frame.pack(anchor=tk.CENTER, pady=5)
        ttk.Label(iter_frame, text="每轮迭代次数:", font=("SimHei", 10)).pack(side=tk.LEFT, padx=5)
        gen_spin = ttk.Spinbox(iter_frame, from_=10, to=80, textvariable=self.raga_n_gen, width=8)
        gen_spin.pack(side=tk.LEFT, padx=5)

        ttk.Label(iter_frame, text="窗口半径（0为自动计算）:", font=("SimHei", 10)).pack(side=tk.LEFT, padx=5)
        radius_spin = ttk.Spinbox(iter_frame, from_=0.0, to=1.0, increment=0.01, textvariable=self.raga_window_radius,
                                  width=8)
        radius_spin.pack(side=tk.LEFT, padx=5)

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
        """加载Excel数据文件（增加数据类型校验）"""
        from tkinter import filedialog
        file_path = filedialog.askopenfilename(
            title="选择数据文件",
            filetypes=[("Excel文件", "*.xlsx;*.xls")]
        )
        if not file_path:
            return

        try:
            # 核心修改：完整加载原始数据，保留所有列
            self.df = pd.read_excel(file_path, dtype=str).convert_dtypes()
            # 分离数值列和非数值列，仅处理数值列用于计算
            numeric_cols = []
            for col in self.df.columns:
                try:
                    # 尝试转换为数值型，保留原始列
                    self.df[col] = pd.to_numeric(self.df[col], errors='ignore')
                    if pd.api.types.is_numeric_dtype(self.df[col]):
                        numeric_cols.append(col)
                except:
                    pass

            self.all_columns = numeric_cols  # 仅数值列用于计算
            self.original_filename = os.path.splitext(os.path.basename(file_path))[0]
            self.file_path_var.set(os.path.basename(file_path))
            self.status_var.set(
                f"数据加载成功：{len(self.df)}条记录，{len(self.df.columns)}列（其中数值列{len(numeric_cols)}个）")
            messagebox.showinfo("成功", f"数据加载完成（保留所有原始列），请点击'开始配置'选择计算用的数值指标列")
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
            label = ttk.Label(main_frame, text="请选择用于计算的数值指标列（至少1个）:",
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

                if method in ["pca", "raga_ppc"]:
                    if method == "raga_ppc":
                        info_text = f"配置完成！\n指标列：{', '.join(selected)}\n注：RAGA-PPC使用固定参数：种群{POPULATION_SIZE} | 加速{ACCELERATE_TIMES}次"
                    else:
                        info_text = f"配置完成！\n指标列：{', '.join(selected)}"
                    messagebox.showinfo("完成", info_text)
                    self.status_var.set(f"配置完成（{self._get_method_name()}）：{len(self.index_columns)}个指标列")
                    self._show_config_info()
                else:
                    messagebox.showinfo("完成", f"已选择指标列：{', '.join(selected)}\n即将进入步骤2：选择负指标")
                    select_negative_indicators()

            btn_frame = ttk.Frame(main_frame)
            btn_frame.pack(fill=tk.X, padx=10, pady=10)

            ttk.Button(btn_frame, text="全选", command=lambda: listbox.selection_set(0, tk.END),
                       style="config.TButton").pack(side=tk.LEFT,
                                                    padx=5)
            ttk.Button(btn_frame, text="取消全选", command=lambda: listbox.selection_clear(0, tk.END),
                       style="config.TButton").pack(
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

            ttk.Button(btn_frame, text="全选", command=lambda: listbox.selection_set(0, tk.END),
                       style="config.TButton").pack(side=tk.LEFT,
                                                    padx=5)
            ttk.Button(btn_frame, text="取消全选", command=lambda: listbox.selection_clear(0, tk.END),
                       style="config.TButton").pack(
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
        if method in ["pca", "raga_ppc"]:
            # 主成分分析法/RAGA-PPC的配置信息
            columns = ("指标名称", "指标类型")
            self.config_tree["columns"] = columns
            for col in columns:
                self.config_tree.heading(col, text=col)
                self.config_tree.column(col, width=180, anchor="center")

            # 填充数据
            for col in self.index_columns:
                if method == "raga_ppc":
                    indicator_type = "投影寻踪指标"
                else:
                    indicator_type = "分析指标"
                self.config_tree.insert("", tk.END, values=(col, indicator_type))
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
            "topsis": "基于熵权法的TOPSIS",
            "raga_ppc": "RAGA-PPC投影寻踪模型"
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

        canvas.draw()
        self.current_figure = fig

    def _show_raga_projection_weights(self, indicators, weights):
        """显示RAGA-PPC投影权重分布"""
        self._clear_visual_area()

        title_label = ttk.Label(
            self.result_visual_frame,
            text="RAGA-PPC投影方向权重分布",
            font=("SimHei", 11, "bold")
        )
        title_label.pack(anchor=tk.CENTER, pady=8)
        title_label.config(foreground="black")

        table_container = ttk.Frame(self.result_visual_frame)
        table_container.pack(expand=True, fill=tk.BOTH, padx=10, pady=5)

        scrollbar_y = ttk.Scrollbar(table_container, orient="vertical")
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)

        tree = ttk.Treeview(
            table_container,
            show="headings",
            yscrollcommand=scrollbar_y.set,
            height=10
        )
        tree.pack(expand=True, fill=tk.BOTH, padx=5, pady=5)
        scrollbar_y.config(command=tree.yview)

        columns = ["指标名称", "投影权重", "影响力占比(%)"]
        tree["columns"] = columns
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=150, anchor="center")

        weights_squared = np.array(weights) ** 2
        influence_ratio = weights_squared / np.sum(weights_squared) * 100

        for idx, (indicator, weight, ratio) in enumerate(zip(indicators, weights, influence_ratio)):
            tree.insert("", tk.END, values=(indicator, f"{weight:.6f}", f"{ratio:.2f}"))

        chart_container = ttk.Frame(self.result_visual_frame)
        chart_container.pack(expand=True, fill=tk.BOTH, padx=10, pady=5)

        fig = plt.figure(figsize=(5, 3), dpi=90)
        ax = fig.add_subplot(111)
        ax.set_facecolor('#f9f9f9')

        y_pos = np.arange(len(indicators))
        ax.barh(y_pos, weights, color='#2196F3', alpha=0.7)
        ax.set_yticks(y_pos)
        ax.set_yticklabels(indicators, fontsize=8)
        ax.set_xlabel('投影权重', fontsize=10)
        ax.set_title(f'RAGA-PPC投影方向权重（固定参数：种群{POPULATION_SIZE} | 加速{ACCELERATE_TIMES}次）', fontsize=10)
        ax.grid(axis='x', linestyle='--', alpha=0.7)

        plt.tight_layout()
        canvas = FigureCanvasTkAgg(fig, chart_container)
        canvas_widget = canvas.get_tk_widget()
        canvas_widget.pack(expand=True, fill=tk.BOTH, padx=5, pady=5)
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

    def _create_log_window(self):
        """创建RAGA-PPC计算日志窗口（同步全局日志缓存）"""
        if self.log_window and self.log_window.winfo_exists():
            self.log_window.lift()
            return self.log_text_widget

        self.log_window = tk.Toplevel(self.root)
        self.log_window.title("RAGA-PPC 计算日志")
        self.log_window.geometry("800x650")
        self.log_window.resizable(True, True)
        self.log_window.configure(bg="#f5f5f5")

        # 绑定日志窗口关闭事件
        def close_log_win():
            self.log_window.destroy()
            self.log_window = None
            self.root.lift()
            self.root.focus_force()

        self.log_window.protocol("WM_DELETE_WINDOW", close_log_win)

        self._center_window(self.log_window, 800, 650)

        log_frame = ttk.Frame(self.log_window, padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(log_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        text_frame = tk.Frame(log_frame, bg="white")
        text_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text_widget = tk.Text(text_frame,
                                       font=("Consolas", 9),
                                       bg="white",
                                       fg="black",
                                       yscrollcommand=scrollbar.set,
                                       wrap=tk.WORD,
                                       padx=5,
                                       pady=5)
        self.log_text_widget.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.log_text_widget.yview)

        btn_frame = ttk.Frame(self.log_window, padding=5)
        btn_frame.pack(fill=tk.X)

        ttk.Button(btn_frame, text="清空日志",
                   command=lambda: self.log_text_widget.delete(1.0, tk.END),
                   style="config.TButton").pack(side=tk.RIGHT, padx=5)

        self.log_window.update_idletasks()
        self.log_text_widget.update_idletasks()

        # 同步全局日志缓存到窗口
        with LOG_LOCK:
            if GLOBAL_LOG_CACHE:
                self.log_text_widget.insert(tk.END, "\n".join(GLOBAL_LOG_CACHE) + "\n")
                self.log_text_widget.see(tk.END)

        log(f"日志窗口已创建，开始记录RAGA-PPC计算过程...", self.log_text_widget)
        log(f"当前固定参数：种群规模={POPULATION_SIZE}, 交叉概率={CROSSOVER_RATE}, 变异概率={MUTATION_RATE}, 优秀个体数={ELITE_INDIVIDUALS}, 加速次数={ACCELERATE_TIMES}",
            self.log_text_widget)

        return self.log_text_widget

    def _stop_calculate(self):
        """停止RAGA-PPC计算"""
        self.stop_flag.set()
        if self.raga_engine:
            self.raga_engine.stop_calculation()
            log("用户终止计算请求已接收", self.log_text_widget)
        self.stop_button.config(state=tk.DISABLED)
        self.calc_button.config(state=tk.NORMAL)
        self.is_calculating = False
        self.status_var.set("计算已终止（用户操作）")

    def _calculate_raga_ppc(self, df_clean, progress_bar, progress_win):
        """RAGA-PPC核心计算逻辑（修复进度条组件销毁问题）"""
        try:
            self.root.after(0, self._create_log_window)
            time.sleep(0.5)

            df_clean = df_clean.fillna(df_clean.mean())
            log("缺失值处理完成：使用均值填充", self.log_text_widget)

            n_gen = self.raga_n_gen.get()
            window_radius = self.raga_window_radius.get() if self.raga_window_radius.get() > 0 else None

            def progress_callback(progress):
                # 修复：先检查进度窗口和进度条是否存在
                if progress_win and progress_win.winfo_exists() and progress_bar and progress_bar.winfo_exists():
                    try:
                        self.root.after(0, lambda: progress_bar.config(value=progress))
                        self.root.after(0, progress_win.update)
                    except Exception as e:
                        log(f"进度更新失败：{str(e)}", self.log_text_widget)

            self.raga_engine = RAGA_PPC_Engine(df_clean, progress_callback=progress_callback,
                                               log_text_widget=self.log_text_widget)

            # 调用solve方法（返回的是归一化到0-1的得分）
            best_vector, normalized_scores = self.raga_engine.solve(
                n_gen=n_gen,
                window_radius=window_radius
            )

            self.results['indicator'] = pd.DataFrame({
                '指标名称': self.index_columns,
                '投影权重': best_vector,
                '影响力占比': (np.array(best_vector) ** 2) / np.sum(np.array(best_vector) ** 2) * 100
            })
            self.results['best_vector'] = best_vector
            self.results['scores'] = normalized_scores  # 直接存储归一化后的得分

            log(f"RAGA-PPC计算完成，得分已归一化到0-1区间", self.log_text_widget)

            # 计算完成回调（先检查窗口是否存在）
            if progress_win and progress_win.winfo_exists():
                self.root.after(0, self._raga_calculate_complete, progress_win)
            else:
                self._raga_calculate_complete(None)

        except Exception as e:
            log(f"RAGA-PPC计算异常: {str(e)}\n{traceback.format_exc()}", self.log_text_widget)
            # 修复：检查进度窗口是否存在
            if progress_win and progress_win.winfo_exists():
                self.root.after(0, self._raga_calculate_error, e, progress_win)
            else:
                self._raga_calculate_error(e, None)
        finally:
            # 修复：最终更新进度条前先检查组件是否存在
            if progress_bar and progress_win and progress_win.winfo_exists() and progress_bar.winfo_exists():
                try:
                    self.root.after(0, lambda: progress_bar.config(value=100))
                    self.root.after(0, progress_win.update)
                except:
                    pass

    def _calculate_raga_ppc_thread(self, df_clean, progress_bar, progress_win):
        """RAGA-PPC线程执行"""
        try:
            self._calculate_raga_ppc(df_clean, progress_bar, progress_win)
        except Exception as e:
            log(f"计算线程异常: {str(e)}", self.log_text_widget)
            if progress_win and progress_win.winfo_exists():
                self.root.after(0, self._raga_calculate_error, e, progress_win)
        finally:
            self.raga_thread = None

    def _raga_calculate_complete(self, progress_win):
        """RAGA-PPC计算完成回调（增加容错）"""
        self.is_calculating = False
        self.calc_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)

        try:
            if progress_win and progress_win.winfo_exists():
                progress_win.destroy()
        except:
            pass

        # 展示结果
        self._show_results()
        self._show_raga_projection_weights(
            self.index_columns,
            self.results['best_vector']
        )
        self.status_var.set(f"RAGA-PPC计算完成（固定参数版）：共{len(self.results['scores'])}个样本得分（0-1区间）")
        messagebox.showinfo("成功", "RAGA-PPC投影寻踪模型计算完成！得分已归一化到0-1区间")

    def _raga_calculate_error(self, error, progress_win):
        """RAGA-PPC计算错误回调（增加容错）"""
        self.is_calculating = False
        self.calc_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)

        try:
            if progress_win and progress_win.winfo_exists():
                progress_win.destroy()
        except:
            pass

        error_msg = f"RAGA-PPC计算出错：{str(error)}"
        log(error_msg, self.log_text_widget)
        messagebox.showerror("计算错误", error_msg)
        self.status_var.set(f"计算失败：{str(error)[:50]}...")

    def _calculate(self):
        """根据选择的方法执行计算（增加进度窗口容错）"""
        if self.df is None:
            messagebox.showwarning("提示", "请先选择数据文件")
            return
        if not self.index_columns:
            messagebox.showwarning("提示", "请先完成参数配置")
            return
        if self.is_calculating:
            messagebox.showinfo("提示", "计算正在进行中，请稍候")
            return

        self.is_calculating = True
        self.stop_flag.clear()
        self.calc_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL if self.calc_method.get() == "raga_ppc" else tk.DISABLED)

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

            # 数据预处理（仅提取计算用的数值列）
            df_clean = self.df[self.index_columns].copy()
            # 转换为数值型（防止原始数据中有文本）
            for col in df_clean.columns:
                df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce')

            # 处理缺失值
            for col in self.index_columns:
                if df_clean[col].isnull().any():
                    df_clean[col].fillna(df_clean[col].mean(), inplace=True)
                    self.status_var.set(f"警告：指标'{col}'存在缺失值，已用均值填充")

            progress_bar['value'] = 30
            progress_win.update()

            # 根据方法选择计算逻辑
            if method == "entropy":
                # ========== 熵值法计算逻辑 ==========
                progress_bar['value'] = 40
                progress_win.update()

                # 数据标准化
                df_standard = df_clean.copy()
                for col in df_standard.columns:
                    if col in self.negative_indicators:
                        # 负指标标准化
                        df_standard[col] = (df_standard[col].max() - df_standard[col]) / (
                                df_standard[col].max() - df_standard[col].min())
                    else:
                        # 正指标标准化
                        df_standard[col] = (df_standard[col] - df_standard[col].min()) / (
                                df_standard[col].max() - df_standard[col].min())

                # 处理标准化后0值
                df_standard = df_standard.replace(0, 1e-10)

                # 计算比重
                df_p = df_standard / df_standard.sum()

                # 计算熵值
                df_ln = np.log(df_p)
                df_ln = df_ln.replace(-np.inf, 0)
                e = -1 / np.log(len(df_clean)) * (df_p * df_ln).sum()

                # 计算差异系数
                d = 1 - e

                # 计算权重
                weights = d / d.sum()

                # 计算综合得分
                scores = (df_standard * weights).sum(axis=1)
                # 归一化到0-1
                scores = (scores - scores.min()) / (scores.max() - scores.min())

                self.results['weights'] = weights
                self.results['scores'] = scores
                self.results['indicator'] = pd.DataFrame({
                    '指标名称': self.index_columns,
                    '熵值': e,
                    '差异系数': d,
                    '权重': weights
                })

                progress_bar['value'] = 90
                progress_win.update()

                # 关闭进度窗口
                progress_win.destroy()

                # 展示结果
                self._show_results()
                self._show_weight_radar_chart(self.index_columns, weights)
                self.status_var.set(f"熵值法计算完成：{len(scores)}个样本得分（0-1区间）")
                messagebox.showinfo("成功", "熵值法计算完成！")

            elif method == "pca":
                # ========== 主成分分析法计算逻辑 ==========
                progress_bar['value'] = 40
                progress_win.update()

                # 数据标准化
                scaler = StandardScaler()
                df_scaled = scaler.fit_transform(df_clean)

                # 主成分分析
                pca = PCA()
                pca_result = pca.fit_transform(df_scaled)

                # 计算贡献率和累计贡献率
                explained_variance = pca.explained_variance_ratio_
                cumulative_variance = np.cumsum(explained_variance)

                # 确定主成分个数（累计贡献率≥85%）
                n_components = np.argmax(cumulative_variance >= 0.85) + 1
                if n_components == 0:
                    n_components = 1

                # 计算综合得分
                pca_selected = pca_result[:, :n_components]
                weights = explained_variance[:n_components] / explained_variance[:n_components].sum()
                scores = np.dot(pca_selected, weights)
                # 归一化到0-1
                scores = (scores - scores.min()) / (scores.max() - scores.min())

                # 存储结果
                self.results['pca'] = pca
                self.results['scaler'] = scaler
                self.results['explained_variance'] = explained_variance
                self.results['cumulative_variance'] = cumulative_variance
                self.results['n_components'] = n_components
                self.results['scores'] = scores

                # 载荷矩阵
                components_df = pd.DataFrame(
                    pca.components_[:n_components],
                    columns=self.index_columns
                )
                self.results['components'] = components_df

                progress_bar['value'] = 90
                progress_win.update()

                # 关闭进度窗口
                progress_win.destroy()

                # 展示结果
                self._show_results()
                self._show_pca_loadings(components_df)
                self.status_var.set(
                    f"主成分分析完成：选择{n_components}个主成分（累计贡献率{cumulative_variance[n_components - 1]:.2%}）")
                messagebox.showinfo("成功",
                                    f"主成分分析法计算完成！\n选择了{n_components}个主成分（累计贡献率{cumulative_variance[n_components - 1]:.2%}）")

            elif method == "topsis":
                # ========== TOPSIS（熵权法）计算逻辑 ==========
                progress_bar['value'] = 40
                progress_win.update()

                # 数据标准化
                df_standard = df_clean.copy()
                for col in df_standard.columns:
                    if col in self.negative_indicators:
                        df_standard[col] = (df_standard[col].max() - df_standard[col]) / (
                                df_standard[col].max() - df_standard[col].min())
                    else:
                        df_standard[col] = (df_standard[col] - df_standard[col].min()) / (
                                df_standard[col].max() - df_standard[col].min())

                # 熵权法计算权重
                df_standard = df_standard.replace(0, 1e-10)
                df_p = df_standard / df_standard.sum()
                df_ln = np.log(df_p)
                df_ln = df_ln.replace(-np.inf, 0)
                e = -1 / np.log(len(df_clean)) * (df_p * df_ln).sum()
                d = 1 - e
                weights = d / d.sum()

                # 加权标准化矩阵
                df_weighted = df_standard * weights

                # 正负理想解
                positive_ideal = df_weighted.max()
                negative_ideal = df_weighted.min()

                # 距离计算
                distance_positive = np.sqrt(((df_weighted - positive_ideal) ** 2).sum(axis=1))
                distance_negative = np.sqrt(((df_weighted - negative_ideal) ** 2).sum(axis=1))

                # 贴近度（TOPSIS得分）
                scores = distance_negative / (distance_positive + distance_negative)
                # 归一化到0-1（本身已在0-1区间，此处为兼容统一逻辑）
                scores = (scores - scores.min()) / (scores.max() - scores.min())

                self.results['weights'] = weights
                self.results['positive_ideal'] = positive_ideal
                self.results['negative_ideal'] = negative_ideal
                self.results['scores'] = scores
                self.results['indicator'] = pd.DataFrame({
                    '指标名称': self.index_columns,
                    '熵值': e,
                    '差异系数': d,
                    '权重': weights
                })

                progress_bar['value'] = 90
                progress_win.update()

                # 关闭进度窗口
                progress_win.destroy()

                # 展示结果
                self._show_results()
                self._show_weight_radar_chart(self.index_columns, weights)
                self.status_var.set(f"TOPSIS计算完成：{len(scores)}个样本得分（0-1区间）")
                messagebox.showinfo("成功", "基于熵权法的TOPSIS计算完成！")

            elif method == "raga_ppc":
                # ========== RAGA-PPC计算逻辑 ==========
                progress_bar['value'] = 20
                progress_win.update()

                # 启动RAGA-PPC线程计算
                self.raga_thread = threading.Thread(
                    target=self._calculate_raga_ppc_thread,
                    args=(df_clean, progress_bar, progress_win),
                    daemon=True
                )
                self.raga_thread.start()

        except Exception as e:
            self.is_calculating = False
            self.calc_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)

            # 容错：检查进度窗口是否存在
            try:
                if 'progress_win' in locals() and progress_win and progress_win.winfo_exists():
                    progress_win.destroy()
            except:
                pass

            error_msg = f"计算出错：{str(e)}"
            log(error_msg, self.log_text_widget if hasattr(self, 'log_text_widget') else None)
            messagebox.showerror("错误", error_msg)
            self.status_var.set(f"计算失败：{str(e)[:50]}...")

    def _show_results(self):
        """展示计算结果表格（核心：保留原始所有列 + 新增综合得分列）"""
        # 清空结果表格
        for item in self.record_tree.get_children():
            self.record_tree.delete(item)
        self.record_tree["columns"] = ()

        if 'scores' not in self.results:
            return

        # ========== 核心修改2：保留原始所有列 + 新增综合得分列 ==========
        # 复制原始完整数据
        result_df = self.df.copy()
        # 新增综合得分列（已归一化到0-1）
        result_df['RAGA-PPC综合得分' if self.calc_method.get() == 'raga_ppc' else '综合得分'] = self.results['scores']

        # 设置表格列（原始所有列 + 综合得分列）
        columns = result_df.columns.tolist()
        self.record_tree["columns"] = columns

        # 配置列属性
        for col in columns:
            self.record_tree.heading(col, text=col)
            # 数值列宽度适配，文本列宽度更大
            if col in ['RAGA-PPC综合得分', '综合得分']:
                self.record_tree.column(col, width=120, anchor="center")
            elif pd.api.types.is_numeric_dtype(result_df[col]):
                self.record_tree.column(col, width=100, anchor="center")
            else:
                self.record_tree.column(col, width=150, anchor="center")

        # 填充数据
        for idx, row in result_df.iterrows():
            values = []
            for col in columns:
                val = row[col]
                # 数值格式化
                if pd.api.types.is_numeric_dtype(result_df[col]):
                    if col in ['RAGA-PPC综合得分', '综合得分']:
                        values.append(f"{val:.6f}")
                    else:
                        values.append(f"{val:.4f}")
                else:
                    values.append(str(val) if pd.notna(val) else "")
            self.record_tree.insert("", tk.END, values=values)

        # 滚动到第一行
        self.record_tree.yview_moveto(0)

    def _save_results(self):
        """保存结果到Excel（核心：原始所有列 + 综合得分列）"""
        if 'scores' not in self.results:
            messagebox.showwarning("提示", "暂无计算结果可保存")
            return

        from tkinter import filedialog
        # 默认文件名
        default_filename = f"{self.original_filename}_{self._get_method_name()}结果.xlsx" if self.original_filename else "计算结果.xlsx"

        file_path = filedialog.asksaveasfilename(
            title="保存结果",
            defaultextension=".xlsx",
            initialfile=default_filename,
            filetypes=[("Excel文件", "*.xlsx")]
        )
        if not file_path:
            return

        try:
            # ========== 核心：保存原始所有列 + 综合得分列 ==========
            # 复制原始完整数据
            save_df = self.df.copy()
            # 新增综合得分列
            score_col_name = 'RAGA-PPC综合得分' if self.calc_method.get() == 'raga_ppc' else '综合得分'
            save_df[score_col_name] = self.results['scores']

            # 创建Excel写入器
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # 主结果表（原始列 + 综合得分）
                save_df.to_excel(writer, sheet_name='综合得分结果', index=False)

                # 指标权重表（如果有）
                if 'indicator' in self.results:
                    self.results['indicator'].to_excel(writer, sheet_name='指标权重', index=False)

                # PCA额外信息（如果有）
                if self.calc_method.get() == 'pca' and 'components' in self.results:
                    self.results['components'].to_excel(writer, sheet_name='主成分载荷矩阵', index=False)
                    # 贡献率信息
                    variance_df = pd.DataFrame({
                        '主成分': [f'主成分{i + 1}' for i in range(len(self.results['explained_variance']))],
                        '贡献率': self.results['explained_variance'],
                        '累计贡献率': self.results['cumulative_variance']
                    })
                    variance_df.to_excel(writer, sheet_name='贡献率', index=False)

            messagebox.showinfo("成功",
                                f"结果已保存到：\n{file_path}\n\n包含：\n1. 原始所有列 + 综合得分列\n2. 指标权重/主成分分析等辅助信息")
            self.status_var.set(f"结果已保存：{os.path.basename(file_path)}")

        except Exception as e:
            messagebox.showerror("错误", f"保存失败：{str(e)}")
            self.status_var.set(f"保存失败：{str(e)[:50]}...")


if __name__ == "__main__":
    # 清理全局线程池缓存
    try:
        GLOBAL_THREAD_POOL.shutdown(wait=False)
    except:
        pass

    root = tk.Tk()
    app = MultiAnalysisTool(root)


    # 主窗口关闭时的清理
    def on_closing():
        try:
            # 停止计算
            if app.is_calculating:
                app._stop_calculate()
            # 关闭线程池
            GLOBAL_THREAD_POOL.shutdown(wait=False)
        except:
            pass
        root.destroy()


    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()