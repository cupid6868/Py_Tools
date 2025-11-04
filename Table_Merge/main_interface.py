import tkinter as tk
from tkinter import messagebox, ttk, font, filedialog
import os
import sys
# 导入Excel合并工具类（关键：从模块中直接导入，实现单EXE打包）
from excel_merger import ExcelMergerPro


class MainInterface:
    def __init__(self, root):
        # 主窗口配置
        self.root = root
        self.root.title("数据处理工具箱 v1.0")
        self.root.geometry("900x650")
        self.root.minsize(800, 600)
        self.root.configure(bg="#f5f5f5")

        # 确保中文显示正常
        self.setup_fonts()

        # 存储工具状态
        self.tools = {
            "excel_merger": {"name": "Excel智能合并工具", "status": "已实现", "color": "#3498db"},
            "data_cleaner": {"name": "数据清洗工具", "status": "开发中", "color": "#95a5a6"},
            "stats_analyzer": {"name": "数据统计分析工具", "status": "开发中", "color": "#95a5a6"},
            "format_converter": {"name": "文件格式转换工具", "status": "开发中", "color": "#95a5a6"},
            "batch_processor": {"name": "批量处理工具", "status": "开发中", "color": "#95a5a6"},
            "visualizer": {"name": "数据可视化工具", "status": "开发中", "color": "#95a5a6"}
        }

        # 创建界面组件
        self.create_widgets()

        # 绑定窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_fonts(self):
        """设置中文字体，确保显示正常"""
        self.title_font = font.Font(family="SimHei", size=20, weight="bold")
        self.subtitle_font = font.Font(family="SimHei", size=12)
        self.btn_font = font.Font(family="SimHei", size=12)
        self.status_font = font.Font(family="SimHei", size=10)
        self.root.option_add("*Font", "SimHei 10")

    def create_widgets(self):
        """创建所有界面组件"""
        # 顶部标题区域
        header_frame = tk.Frame(self.root, bg="#f5f5f5", pady=20)
        header_frame.pack(fill=tk.X, padx=20)

        tk.Label(
            header_frame,
            text="数据处理工具箱",
            font=self.title_font,
            bg="#f5f5f5",
            fg="#2c3e50"
        ).pack(anchor=tk.W)

        tk.Label(
            header_frame,
            text="高效处理各类数据任务，简化您的工作流程",
            font=self.subtitle_font,
            bg="#f5f5f5",
            fg="#666"
        ).pack(anchor=tk.W, pady=5)

        # 功能按钮区域（使用网格布局更整齐）
        tools_frame = tk.Frame(self.root, bg="#f5f5f5", padx=40, pady=10)
        tools_frame.pack(fill=tk.BOTH, expand=True)

        # 按钮样式配置
        btn_style = {
            "font": self.btn_font,
            "width": 22,
            "height": 3,
            "bd": 0,
            "relief": tk.RAISED,
            "cursor": "hand2",
            "activebackground": "#2980b9"
        }

        # 第一行工具按钮
        self.excel_merge_btn = tk.Button(
            tools_frame,
            text=self.tools["excel_merger"]["name"],
            bg=self.tools["excel_merger"]["color"],
            fg="white",
            command=self.launch_excel_merger, **btn_style
        )
        self.excel_merge_btn.grid(row=0, column=0, padx=15, pady=20)

        self.data_clean_btn = tk.Button(
            tools_frame,
            text=self.tools["data_cleaner"]["name"],
            bg=self.tools["data_cleaner"]["color"],
            fg="white",
            command=lambda: self.show_under_development("数据清洗工具"),
            **btn_style
        )
        self.data_clean_btn.grid(row=0, column=1, padx=15, pady=20)

        # 第二行工具按钮
        self.stats_btn = tk.Button(
            tools_frame,
            text=self.tools["stats_analyzer"]["name"],
            bg=self.tools["stats_analyzer"]["color"],
            fg="white",
            command=lambda: self.show_under_development("数据统计分析工具"), **btn_style
        )
        self.stats_btn.grid(row=1, column=0, padx=15, pady=20)

        self.convert_btn = tk.Button(
            tools_frame,
            text=self.tools["format_converter"]["name"],
            bg=self.tools["format_converter"]["color"],
            fg="white",
            command=lambda: self.show_under_development("文件格式转换工具"),
            **btn_style
        )
        self.convert_btn.grid(row=1, column=1, padx=15, pady=20)

        # 第三行工具按钮
        self.batch_btn = tk.Button(
            tools_frame,
            text=self.tools["batch_processor"]["name"],
            bg=self.tools["batch_processor"]["color"],
            fg="white",
            command=lambda: self.show_under_development("批量处理工具"), **btn_style
        )
        self.batch_btn.grid(row=2, column=0, padx=15, pady=20)

        self.visual_btn = tk.Button(
            tools_frame,
            text=self.tools["visualizer"]["name"],
            bg=self.tools["visualizer"]["color"],
            fg="white",
            command=lambda: self.show_under_development("数据可视化工具"),
            **btn_style
        )
        self.visual_btn.grid(row=2, column=1, padx=15, pady=20)

        # 底部状态区域
        footer_frame = tk.Frame(self.root, bg="#e0e0e0", height=50)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)

        # 状态标签
        self.status_var = tk.StringVar(value="就绪 - 请选择需要使用的工具")
        tk.Label(
            footer_frame,
            textvariable=self.status_var,
            font=self.status_font,
            bg="#e0e0e0",
            fg="#333"
        ).pack(side=tk.LEFT, padx=20, pady=15)

        # 版本信息
        tk.Label(
            footer_frame,
            text="v1.0 | 数据处理工具箱",
            font=self.status_font,
            bg="#e0e0e0",
            fg="#666"
        ).pack(side=tk.RIGHT, padx=20, pady=15)

    def launch_excel_merger(self):
        """启动Excel合并工具（作为子窗口）"""
        try:
            self.status_var.set(f"启动 {self.tools['excel_merger']['name']}...")
            self.root.update()  # 刷新界面显示状态

            # 创建子窗口并实例化合并工具
            merger_window = tk.Toplevel(self.root)
            merger_window.title(self.tools["excel_merger"]["name"])
            merger_window.geometry("1100x700")
            merger_window.minsize(1000, 650)
            merger_window.option_add("*Font", "SimHei 10")

            # 子窗口关闭时更新状态
            def on_merger_close():
                merger_window.destroy()
                self.status_var.set("就绪 - 已关闭Excel智能合并工具")

            merger_window.protocol("WM_DELETE_WINDOW", on_merger_close)

            # 实例化合并工具
            ExcelMergerPro(merger_window)
            self.status_var.set(f"运行中 - {self.tools['excel_merger']['name']}")

        except Exception as e:
            self.status_var.set("错误 - 启动工具失败")
            messagebox.showerror("启动失败", f"无法启动Excel智能合并工具：\n{str(e)}")

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
    # 解决高DPI屏幕显示问题
    try:
        from ctypes import windll

        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass

    root = tk.Tk()
    app = MainInterface(root)
    root.mainloop()