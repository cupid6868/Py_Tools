import tkinter as tk
from tkinter import messagebox, ttk, font, filedialog
import os
import sys
import time
import threading
from queue import Queue
import math
from typing import List, Optional
import glob
import pandas as pd

# 导入 Excel合并工具类（如果不存在可注释，不影响格式转换功能）
try:
    from excel_merge_multi import ExcelMergerPro
except ImportError:
    class ExcelMergerPro:
        def __init__(self, root):
            messagebox.showwarning("提示", "Excel合并工具类未找到，该功能暂时无法使用")

# 导入 综合分析工具类（如果不存在可注释，不影响格式转换功能）
try:
    from Comprehensive_index import MultiAnalysisTool
except ImportError:
    class MultiAnalysisTool:
        def __init__(self, root):
            messagebox.showwarning("提示", "综合分析工具类未找到，该功能暂时无法使用")

# 导入 文件格式转换工具类（DTA转CSV核心功能），已移至外部文件 Diff_Format_Convert.py
try:
    from Diff_Format_Convert import DtaToCsvConverter
except ImportError:
    # 定义一个占位类，当外部文件不存在时，提示用户该功能不可用
    class DtaToCsvConverter:
        def __init__(self, root):
            messagebox.showwarning("提示", "DTA转CSV转换工具类未找到或导入失败，请确保 Diff_Format_Convert.py 文件在同一目录下！")

# ========== 新增：导入单表合并工具类 ==========
try:
    from Merge_Single import ExcelMergerPro_2  # 从Merge_Single模块导入ExcelMergerPro_2
except ImportError:
    class ExcelMergerPro_2:
        def __init__(self, root):
            messagebox.showwarning("提示", "单表合并工具类未找到，请确保 Merge_Single.py 文件在同一目录下！")

# ========== 主界面类 ==========
class MainInterface:
    def __init__(self, root):
        # 主窗口配置
        self.root = root
        self.root.title("数据处理工具箱 v2.0")
        self.root.geometry("900x650")
        self.root.minsize(800, 600)
        self.root.configure(bg="#f8f9fa")

        # 关键设置：先隐藏主窗口且不绘制内容
        self.root.withdraw()
        self.root.update()

        # 确保中文显示正常
        self.setup_fonts()

        # 主题颜色配置
        self.colors = {
            "primary": "#3f72af",
            "primary_light": "#5b8fb9",
            "primary_dark": "#2c5282",
            "secondary": "#b7adcf",
            "success": "#40a860",
            "warning": "#ffb627",
            "info": "#38b2ac",
            "light": "#ffffff",
            "dark": "#2d3748",
            "gray": "#a0aec0",
            "light_gray": "#e2e8f0",
            "background": "#f8f9fa"
        }

        # 存储工具状态（修改：批量处理工具→单表合并工具，标记为已实现）
        self.tools = {
            "excel_merger": {"name": "Excel智能合并工具", "status": "已实现", "color": self.colors["primary"]},
            "data_cleaner": {"name": "数据清洗工具", "status": "开发中", "color": self.colors["gray"]},
            "stats_analyzer": {"name": "数据统计分析工具", "status": "已实现", "color": self.colors["primary"]},
            "format_converter": {"name": "文件格式转换工具", "status": "已实现", "color": self.colors["primary"]},
            "batch_processor": {"name": "单表合并工具", "status": "已实现", "color": self.colors["primary"]},  # 修改此处
            "visualizer": {"name": "数据可视化工具", "status": "开发中", "color": self.colors["gray"]}
        }

        # 绑定窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # 先显示加载窗口，完成后再创建主界面元素
        self.add_loading_animation()

    def setup_fonts(self):
        """设置中文字体，确保显示正常"""
        self.title_font = font.Font(family=["SimHei", "WenQuanYi Micro Hei", "Heiti TC"], size=22, weight="bold")
        self.subtitle_font = font.Font(family=["SimHei", "WenQuanYi Micro Hei", "Heiti TC"], size=12)
        self.btn_font = font.Font(family=["SimHei", "WenQuanYi Micro Hei", "Heiti TC"], size=12)
        self.status_font = font.Font(family=["SimHei", "WenQuanYi Micro Hei", "Heiti TC"], size=10)

        # 全局字体设置
        default_font = font.nametofont("TkDefaultFont")
        default_font.configure(family=["SimHei", "WenQuanYi Micro Hei", "Heiti TC"], size=10)
        default_font.configure()

    def create_widgets(self):
        """创建所有界面组件（在加载完成后调用）"""
        # 顶部标题区域
        header_frame = tk.Frame(self.root, bg=self.colors["light"], pady=20, relief=tk.FLAT, bd=0)
        header_frame.pack(fill=tk.X)

        # 标题区域阴影
        shadow_frame = tk.Frame(self.root, bg=self.colors["light_gray"], height=2)
        shadow_frame.pack(fill=tk.X)

        # 标题内容
        title_content = tk.Frame(header_frame, bg=self.colors["light"])
        title_content.pack(anchor=tk.W, padx=40)

        tk.Label(
            title_content,
            text="数据处理工具箱",
            font=self.title_font,
            bg=self.colors["light"],
            fg=self.colors["dark"]
        ).pack(anchor=tk.W, pady=(0, 5))

        tk.Label(
            title_content,
            text="高效处理各类数据任务，简化您的工作流程",
            font=self.subtitle_font,
            bg=self.colors["light"],
            fg="#4a5568"
        ).pack(anchor=tk.W)

        # 功能按钮区域
        tools_frame = tk.Frame(self.root, bg=self.colors["background"], padx=20, pady=20)
        tools_frame.pack(fill=tk.BOTH, expand=True)

        # 卡片容器
        card_frame = tk.Frame(tools_frame, bg=self.colors["light"], bd=0, relief=tk.RAISED)
        card_frame.pack(fill=tk.BOTH, expand=True, padx=40, pady=10)

        # 卡片阴影
        card_shadow = tk.Frame(tools_frame, bg=self.colors["light_gray"])
        card_shadow.pack(fill=tk.BOTH, expand=True, padx=42, pady=(0, 12))
        card_shadow.lower()

        # 按钮样式
        btn_style = {
            "font": self.btn_font,
            "width": 22,
            "height": 3,
            "bd": 0,
            "relief": tk.FLAT,
            "cursor": "hand2",
            "highlightthickness": 0
        }

        # 第一行按钮
        self.excel_merge_btn = tk.Button(
            card_frame,
            text=self.tools["excel_merger"]["name"],
            bg=self.tools["excel_merger"]["color"],
            fg="white",
            command=self.launch_excel_merger, **btn_style
        )
        self.excel_merge_btn.grid(row=0, column=0, padx=20, pady=20)
        self.add_button_hover_effect(self.excel_merge_btn, self.colors["primary"], self.colors["primary_light"],
                                     self.colors["primary_dark"])

        self.data_clean_btn = tk.Button(
            card_frame,
            text=self.tools["data_cleaner"]["name"],
            bg=self.tools["data_cleaner"]["color"],
            fg="white",
            command=lambda: self.show_under_development("数据清洗工具"),
            **btn_style
        )
        self.data_clean_btn.grid(row=0, column=1, padx=20, pady=20)
        self.add_button_hover_effect(self.data_clean_btn, self.colors["gray"], self.colors["light_gray"],
                                     self.colors["dark"])

        # 第二行按钮
        self.stats_btn = tk.Button(
            card_frame,
            text=self.tools["stats_analyzer"]["name"],
            bg=self.tools["stats_analyzer"]["color"],
            fg="white",
            command=self.launch_stats_analyzer,
            **btn_style
        )
        self.stats_btn.grid(row=1, column=0, padx=20, pady=20)
        self.add_button_hover_effect(self.stats_btn, self.colors["primary"], self.colors["primary_light"],
                                     self.colors["primary_dark"])

        # 文件格式转换按钮（已实现，绑定到转换工具）
        self.convert_btn = tk.Button(
            card_frame,
            text=self.tools["format_converter"]["name"],
            bg=self.tools["format_converter"]["color"],
            fg="white",
            command=self.launch_format_converter,
            **btn_style
        )
        self.convert_btn.grid(row=1, column=1, padx=20, pady=20)
        self.add_button_hover_effect(self.convert_btn, self.colors["primary"], self.colors["primary_light"],
                                     self.colors["primary_dark"])

        # 第三行按钮（修改：单表合并工具，绑定到ExcelMergerPro_2）
        self.batch_btn = tk.Button(
            card_frame,
            text=self.tools["batch_processor"]["name"],
            bg=self.tools["batch_processor"]["color"],
            fg="white",
            command=self.launch_single_merge,  # 修改此处
            **btn_style
        )
        self.batch_btn.grid(row=2, column=0, padx=20, pady=20)
        self.add_button_hover_effect(self.batch_btn, self.colors["primary"], self.colors["primary_light"],  # 修改hover颜色
                                     self.colors["primary_dark"])

        self.visual_btn = tk.Button(
            card_frame,
            text=self.tools["visualizer"]["name"],
            bg=self.tools["visualizer"]["color"],
            fg="white",
            command=lambda: self.show_under_development("数据可视化工具"),
            **btn_style
        )
        self.visual_btn.grid(row=2, column=1, padx=20, pady=20)
        self.add_button_hover_effect(self.visual_btn, self.colors["gray"], self.colors["light_gray"],
                                     self.colors["dark"])

        # 底部状态区域
        footer_frame = tk.Frame(self.root, bg=self.colors["dark"], height=50)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)

        # 状态标签
        self.status_var = tk.StringVar(value="就绪 - 请选择需要使用的工具")
        tk.Label(
            footer_frame,
            textvariable=self.status_var,
            font=self.status_font,
            bg=self.colors["dark"],
            fg="white"
        ).pack(side=tk.LEFT, padx=20, pady=15)

        # 版本信息
        tk.Label(
            footer_frame,
            text="v2.0 | 数据处理工具箱",
            font=self.status_font,
            bg=self.colors["dark"],
            fg="#bbb"
        ).pack(side=tk.RIGHT, padx=20, pady=15)

    # ========== 新增：启动单表合并工具方法 ==========
    def launch_single_merge(self):
        """启动单表合并工具（ExcelMergerPro_2）"""
        try:
            self.status_var.set(f"启动 {self.tools['batch_processor']['name']}...")
            self.root.update()

            self.root.withdraw()
            single_merge_window = tk.Toplevel(self.root)
            single_merge_window.title(self.tools["batch_processor"]["name"])
            single_merge_window.geometry("1100x700")
            single_merge_window.minsize(1000, 650)
            single_merge_window.configure(bg=self.colors["background"])
            self.center_window(single_merge_window)

            def on_single_merge_close():
                single_merge_window.destroy()
                self.status_var.set(f"就绪 - 已关闭{self.tools['batch_processor']['name']}")
                self.root.deiconify()

            single_merge_window.protocol("WM_DELETE_WINDOW", on_single_merge_close)
            ExcelMergerPro_2(single_merge_window)  # 实例化单表合并工具类
            self.status_var.set(f"运行中 - {self.tools['batch_processor']['name']}")

        except Exception as e:
            self.root.deiconify()
            self.status_var.set("错误 - 启动工具失败")
            messagebox.showerror("启动失败", f"无法启动单表合并工具：\n{str(e)}")

    def add_button_hover_effect(self, button, normal_bg, hover_bg, press_bg):
        """为按钮添加悬停和点击效果"""

        def on_enter(e):
            button['background'] = hover_bg

        def on_leave(e):
            button['background'] = normal_bg

        def on_press(e):
            button['background'] = press_bg

        def on_release(e):
            button['background'] = hover_bg

        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)
        button.bind("<ButtonPress-1>", on_press)
        button.bind("<ButtonRelease-1>", on_release)

    def add_loading_animation(self):
        """加载窗口优先显示，主窗口元素在加载完成后才创建"""
        self.loading_window = tk.Toplevel()
        self.loading_window.overrideredirect(True)
        self.loading_window.geometry("300x100")
        self.loading_window.configure(bg=self.colors["light"])
        self.loading_window.attributes("-alpha", 0.95)

        self.loading_window.attributes("-topmost", True)
        self.loading_window.grab_set_global()

        self.center_window(self.loading_window)

        self.loading_window.update_idletasks()
        self.loading_window.lift()

        tk.Label(
            self.loading_window,
            text="加载数据处理工具箱...",
            font=self.subtitle_font,
            bg=self.colors["light"],
            fg=self.colors["dark"]
        ).pack(pady=20)

        style = ttk.Style()
        style.configure("TProgressbar",
                        troughcolor=self.colors["light_gray"],
                        background=self.colors["primary"])

        progress = ttk.Progressbar(self.loading_window,
                                   length=200,
                                   mode='indeterminate',
                                   style="TProgressbar")
        progress.pack()
        progress.start()

        def close_loading():
            progress.stop()
            self.loading_window.grab_release()
            self.loading_window.attributes("-topmost", False)
            self.loading_window.destroy()

            self.create_widgets()

            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()

        self.loading_window.after(1000, close_loading)

    def center_window(self, window):
        """使窗口居中显示"""
        window.update_idletasks()
        width = window.winfo_width()
        height = window.winfo_height()
        x = (window.winfo_screenwidth() // 2) - (width // 2)
        y = (window.winfo_screenheight() // 2) - (height // 2)
        window.geometry(f"{width}x{height}+{x}+{y}")

    def launch_excel_merger(self):
        """启动Excel合并工具"""
        try:
            self.status_var.set(f"启动 {self.tools['excel_merger']['name']}...")
            self.root.update()

            self.root.withdraw()
            merger_window = tk.Toplevel(self.root)
            merger_window.title(self.tools["excel_merger"]["name"])
            merger_window.geometry("1100x700")
            merger_window.minsize(1000, 650)
            merger_window.configure(bg=self.colors["background"])
            self.center_window(merger_window)

            def on_merger_close():
                merger_window.destroy()
                self.status_var.set("就绪 - 已关闭Excel智能合并工具")
                self.root.deiconify()

            merger_window.protocol("WM_DELETE_WINDOW", on_merger_close)
            ExcelMergerPro(merger_window)
            self.status_var.set(f"运行中 - {self.tools['excel_merger']['name']}")

        except Exception as e:
            self.root.deiconify()
            self.status_var.set("错误 - 启动工具失败")
            messagebox.showerror("启动失败", f"无法启动Excel智能合并工具：\n{str(e)}")

    def launch_stats_analyzer(self):
        """启动数据统计分析工具"""
        try:
            self.status_var.set(f"启动 {self.tools['stats_analyzer']['name']}...")
            self.root.update()

            self.root.withdraw()
            stats_window = tk.Toplevel(self.root)
            stats_window.title(self.tools["stats_analyzer"]["name"])
            stats_window.geometry("1000x750")
            stats_window.resizable(False, False)
            stats_window.configure(bg=self.colors["background"])
            self.center_window(stats_window)

            def on_stats_close():
                stats_window.destroy()
                self.status_var.set("就绪 - 已关闭数据统计分析工具")
                self.root.deiconify()

            stats_window.protocol("WM_DELETE_WINDOW", on_stats_close)
            MultiAnalysisTool(stats_window)
            self.status_var.set(f"运行中 - {self.tools['stats_analyzer']['name']}")

        except Exception as e:
            self.root.deiconify()
            self.status_var.set("错误 - 启动工具失败")
            messagebox.showerror("启动失败", f"无法启动数据统计分析工具：\n{str(e)}")

    def launch_format_converter(self):
        """启动文件格式转换工具（DTA转CSV）"""
        try:
            self.status_var.set(f"启动 {self.tools['format_converter']['name']}...")
            self.root.update()

            self.root.withdraw()
            converter_window = tk.Toplevel(self.root)
            converter_window.title(self.tools["format_converter"]["name"] + " - DTA转CSV")
            converter_window.geometry("1200x800")
            converter_window.minsize(1000, 700)
            converter_window.configure(bg=self.colors["background"])
            self.center_window(converter_window)

            def on_converter_close():
                converter_window.destroy()
                self.status_var.set("就绪 - 已关闭文件格式转换工具")
                self.root.deiconify()

            converter_window.protocol("WM_DELETE_WINDOW", on_converter_close)
            # 实例化外部文件中的DTA转CSV工具
            DtaToCsvConverter(converter_window)
            self.status_var.set(f"运行中 - {self.tools['format_converter']['name']}")

        except Exception as e:
            self.root.deiconify()
            self.status_var.set("错误 - 启动工具失败")
            messagebox.showerror("启动失败", f"无法启动文件格式转换工具：\n{str(e)}")

    def show_under_development(self, tool_name):
        """显示待开发提示"""
        self.status_var.set(f"提示 - {tool_name} 正在开发中")
        messagebox.showinfo("功能开发中", f"{tool_name} 正在积极开发中，敬请期待！")

    def on_close(self):
        """窗口关闭确认"""
        if messagebox.askyesno("确认退出", "确定要关闭数据处理工具箱吗？"):
            self.root.destroy()
            sys.exit(0)


if __name__ == "__main__":
    # 高DPI屏幕适配
    try:
        from ctypes import windll

        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass

    root = tk.Tk()
    app = MainInterface(root)
    root.mainloop()