import os
import tkinter as tk
import ttkbootstrap as ttk_bootstrap
import threading
import sys

# 导入配置
from configs.api_config import load_api_configs, save_api_configs
from configs.excel_header_config import get_default_columns, save_custom_columns

# 导入GUI组件 - 修改为从正确模块导入
from gui.ui_setup import setup_ui
from gui.gui_components import init_styles
from gui.gui_actions import create_actions_frame
from gui.gui_components import create_tools_frame

# 导入功能模块
from strategies.extract import extract_content
from strategies.review import generate_review

# 导入工具函数
from utils.excel_utils import (
    select_excel,
    display_excel_info,
    create_excel,
    check_excel_columns,
)
from utils.excel_utils import save_to_excel_with_format, update_excel_info_label

# 修正导入语句，只从pdf_manager导入函数，不再从pdf_analysis导入
from utils.pdf_manager import (
    analyze_papers,
    regenerate,
    process_papers_async,
    perform_regenerate,
    select_pdfs,
    add_pdf,
    remove_pdf,
    display_pdf_info,
)

from gui.ui_utils import (
    disable_analysis_buttons,
    enable_analysis_buttons,
    update_progress_status,
)
from gui.ui_utils import (
    append_response_chunk,
    save_excel_result,
    on_analysis_complete,
    show_error_and_reset,
    cancel_analysis,
)
from utils.app_manager import init_app_services, cleanup_app_services


class PaperAnalyzer:
    def __init__(self, config_path=None):
        self.config_path = config_path
        try:
            self.root = ttk_bootstrap.Window(themename="darkly")
            self.ttk = ttk_bootstrap
            self.root.lift()
            self.root.attributes("-topmost", True)
            self.root.after_idle(self.root.attributes, "-topmost", False)
            self.root.title("论文分析工具")
            self.root.geometry("2000x1200")
            self.root.minsize(2000, 1200)
            # 添加语言切换变量
            self.language_var = tk.StringVar(value="中文")

            # 确保用户数据目录存在
            user_docs = os.path.join(
                os.path.expanduser("~"), "Documents", "论文分析工具"
            )
            if not os.path.exists(user_docs):
                os.makedirs(user_docs)

            # 重置表头为默认（程序启动时始终使用默认表头）
            try:
                default_columns = get_default_columns()
                save_custom_columns(default_columns)
                print(f"表头已重置为默认，共{len(default_columns)}列")
            except Exception as e:
                print(f"重置表头失败: {str(e)}")

            init_app_services(self.root)
            self.root.protocol("WM_DELETE_WINDOW", self.on_close)

            self.colors = {
                "bg": "#2d2d2d",
                "secondary_bg": "#363636",
                "fg": "#f0f0f0",
                "button": "#404040",
                "button_hover": "#505050",
                "button_active": "#606060",
                "frame": "#363636",
                "accent": "#4a90d9",
                "border": "#454545",
                "entry_bg": "#404040",
                "progress": "#4a8cca",
                "gray_text": "#c0c0c0",
                "path_text": "#b5b5b5",
            }

            self.fonts = {
                "title": ("微软雅黑", 13, "bold"),
                "text": ("微软雅黑", 10),
                "button": ("微软雅黑", 11, "bold"),
                "console": ("Consolas", 10),
            }

            self.excel_path = tk.StringVar(value="")

            # 先初始化为空值，然后从配置中加载
            self.api_model_var = tk.StringVar(value="")
            self.api_key_var = tk.StringVar(value="")
            self.api_url_var = tk.StringVar(value="")
            self.remember_key_var = tk.BooleanVar(value=(True))

            # 加载API配置
            self.load_api_configs()

            # 加载自定义表头
            try:
                from configs.excel_header_config import load_custom_columns

                self.custom_columns = load_custom_columns()
                print(f"加载了自定义表头配置: {len(self.custom_columns)}列")
            except Exception as e:
                print(f"加载自定义表头失败: {e}")

            self.active_excel_config = tk.StringVar(value="默认")
            self.analysis_completed = (
                False  # 表示当前论文分析任务尚未完成，任务完成后会设为 True。
            )
            self.cancel_analysis_requested = False  # 是否请求取消论文分析任务
            self.cancel_translation_requested = False  # 添加翻译取消标志

            # 初始化样式
            init_styles(self)

            # 绑定界面函数
            self.setup_ui = lambda: setup_ui(self)
            self.create_actions_frame = lambda parent_frame: create_actions_frame(
                self, parent_frame
            )
            self.create_tools_frame = lambda parent_frame: create_tools_frame(
                self, parent_frame
            )
            self.manage_excel_configs = lambda: self._manage_excel_configs()

            # 绑定Excel操作函数
            self.select_excel = lambda: select_excel(self)
            self.display_excel_info = lambda path: display_excel_info(self, path)
            self.create_excel = lambda: create_excel(self)
            self.check_excel_columns = lambda path: check_excel_columns(self, path)
            self.save_to_excel_with_format = lambda df, path: save_to_excel_with_format(
                df, path
            )
            self.update_excel_info_label = lambda path: update_excel_info_label(
                self, path
            )

            # 绑定PDF分析函数
            self.analyze_papers = lambda: analyze_papers(self)
            self.regenerate = lambda: regenerate(self)
            self.cancel_analysis = lambda: cancel_analysis(self)
            self.process_papers_async = (
                lambda df, url, key, total: process_papers_async(
                    self, df, url, key, total
                )
            )
            self.perform_regenerate = lambda: perform_regenerate(self)

            # 绑定UI工具函数
            self.disable_analysis_buttons = lambda: disable_analysis_buttons(self)
            self.enable_analysis_buttons = lambda: enable_analysis_buttons(self)
            self.update_progress_status = lambda current, total: update_progress_status(
                self, current, total
            )
            self.append_response_chunk = lambda text: append_response_chunk(self, text)
            self.save_excel_result = lambda df: save_excel_result(self, df)
            self.on_analysis_complete = lambda: on_analysis_complete(self)
            self.show_error_and_reset = lambda msg: show_error_and_reset(self, msg)

            # 绑定PDF文件管理函数
            self.select_pdfs = lambda: select_pdfs(self)  # 保持绑定以保证兼容性
            self.display_pdf_info = lambda paths: display_pdf_info(self, paths)
            self.add_pdf = lambda: add_pdf(self)
            self.remove_pdf = lambda: remove_pdf(self)

            # 绑定翻译和提取功能
            self.extract_content = lambda: extract_content(self)
            self.generate_review = lambda: generate_review(self)
            # 删除论文搜索工具绑定行

            # 初始化UI组件
            self.initialize_ui()

        except Exception as e:
            import traceback

            print(f"初始化界面时出错: {str(e)}")
            print(traceback.format_exc())
            raise

    def run_external_tool(self, tool_dir, tool_file):
        """异步运行外部工具，优化后的路径处理逻辑，兼容打包环境"""
        # 由于删除了论文搜索工具，这个方法可以保留但不再使用
        import threading
        import os
        import sys
        import subprocess

        def run_tool():
            try:
                # 调试信息
                print(f"正在尝试启动外部工具: {tool_dir}/{tool_file}")

                # 确定应用程序根目录
                if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
                    # 如果是打包后的应用，使用exe所在目录
                    app_dir = os.path.dirname(sys.executable)
                    print(f"检测到打包环境，应用根目录: {app_dir}")
                else:
                    # 开发环境使用脚本所在目录
                    app_dir = os.path.dirname(os.path.abspath(__file__))
                    # 由于app.py在gui文件夹下，所以需要向上一级
                    app_dir = os.path.dirname(app_dir)
                    print(f"检测到开发环境，应用根目录: {app_dir}")

                # 构建工具目录和文件的完整路径
                full_tool_dir = os.path.join(app_dir, tool_dir)
                full_tool_path = os.path.join(full_tool_dir, tool_file)

                print(f"工具目录: {full_tool_dir}")
                print(f"工具路径: {full_tool_path}")

                # 检查文件是否存在
                if os.path.exists(full_tool_path):
                    print(f"工具文件存在，准备启动")

                    # 获取Python解释器路径
                    if getattr(sys, "frozen", False):
                        # 直接使用系统Python解释器
                        print("使用系统Python解释器")
                        python_exe = "python"
                    else:
                        # 使用当前Python解释器
                        python_exe = sys.executable
                        print(f"使用当前Python解释器: {python_exe}")

                    # 列出目录内容，用于调试
                    print(f"目录 {full_tool_dir} 内容:")
                    for item in os.listdir(full_tool_dir):
                        print(f"  - {item}")

                    # 启动子进程运行外部工具
                    if os.name == "nt":  # Windows
                        # 使用cmd来确保环境变量加载正确
                        cmd = f'cd /d "{full_tool_dir}" && {python_exe} "{tool_file}"'
                        print(f"执行命令: {cmd}")
                        subprocess.Popen(
                            cmd, shell=True, creationflags=subprocess.CREATE_NO_WINDOW
                        )
                    else:  # Linux/Mac
                        cmd = [python_exe, full_tool_path]
                        print(f"执行命令: {' '.join(cmd)}")
                        subprocess.Popen(
                            cmd,
                            cwd=full_tool_dir,
                            stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE,
                        )

                    print("外部工具启动命令已执行")
                else:
                    print(f"错误: 工具文件不存在: {full_tool_path}")

                    # 尝试直接找到并使用ui_interface.py
                    alternative_paths = [
                        os.path.join(app_dir, "paper_find", "ui_interface.py"),
                        os.path.join(
                            os.path.dirname(app_dir), "paper_find", "ui_interface.py"
                        ),
                        # 添加更多可能的路径
                    ]

                    for alt_path in alternative_paths:
                        print(f"尝试替代路径: {alt_path}")
                        if os.path.exists(alt_path):
                            alt_dir = os.path.dirname(alt_path)
                            print(f"找到替代文件，尝试启动: {alt_path}")
                            if os.name == "nt":
                                cmd = f'cd /d "{alt_dir}" && python "{os.path.basename(alt_path)}"'
                                subprocess.Popen(
                                    cmd,
                                    shell=True,
                                    creationflags=subprocess.CREATE_NO_WINDOW,
                                )
                            else:
                                subprocess.Popen(
                                    [sys.executable, alt_path], cwd=alt_dir
                                )
                            return
                    print("无法找到任何可用的UI界面文件")

            except Exception as e:
                import traceback

                print(f"启动工具时出错: {str(e)}")
                print(traceback.format_exc())

        # 使用线程异步启动工具，不阻塞主界面
        threading.Thread(target=run_tool, daemon=True).start()

    def on_close(self):
        """关闭程序前保存API配置"""
        try:
            # 先保存API配置
            self.save_api_configs()
            # 再清理应用服务
            cleanup_app_services()
        finally:
            self.root.destroy()

    def initialize_ui(self):
        try:
            # 调用setup_ui()构建UI组件，并保存output_text引用
            ui_components = self.setup_ui()
            self.ui_components = ui_components
            if "output_text" in ui_components:
                self.output_text = ui_components["output_text"]
            else:
                raise Exception("output_text组件未能正确创建")

            # 调用欢迎信息函数
            self.display_welcome_message()

            # 确保"内容提取"按钮绑定正确的函数
            if hasattr(self, "extract_btn"):
                self.extract_btn.config(command=self.extract_content)

            # 设置分隔条位置
            def ensure_paned_position():
                if (
                    hasattr(self, "ui_components")
                    and "main_paned" in self.ui_components
                ):
                    paned = self.ui_components["main_paned"]
                    width = paned.winfo_width()
                    # 设置左侧宽度为整体宽度的20%
                    paned.sashpos(0, int(width * 0.2))

            self.root.after(200, ensure_paned_position)

            # 添加版本和作者信息到状态栏右侧
            if hasattr(self, "status_bar"):
                version_label = self.ttk.Label(
                    self.root,
                    text="tju-wyh、dny, v1.0",
                    foreground="#888888",
                    padding=(5, 2),
                )
                version_label.place(
                    relx=1.0, rely=1.0, anchor="se", bordermode="outside"
                )

        except Exception as e:
            import traceback

            print(f"初始化UI时出错: {str(e)}")
            print(traceback.format_exc())
            raise

    def toggle_language(self):
        """切换界面语言"""
        if self.language_var.get() == "中文":
            self.language_var.set("English")
        else:
            self.language_var.set("中文")
        self.update_ui_language()

    def update_ui_language(self):
        """更新UI界面上的语言"""
        from gui.ui_utils import update_button_texts

        # 执行更新
        update_button_texts(self)

        # 确保语言切换按钮的文本是正确的
        if hasattr(self, "lang_btn"):
            is_english = self.language_var.get() == "English"
            self.lang_btn.configure(
                text="切换为中文" if is_english else "Switch to English"
            )

        # 重新显示欢迎信息
        self.display_welcome_message()

        # 强制更新窗口
        self.root.update_idletasks()

    def display_welcome_message(self):
        # 确保output_text已初始化
        if not hasattr(self, "output_text"):
            print("错误：output_text未初始化")
            return

        self.output_text.delete(1.0, tk.END)

        is_english = self.language_var.get() == "English"

        if is_english:
            self.output_text.insert(
                tk.END, "=== Welcome to Paper Analysis Tool ===\n\n", "header"
            )
            self.output_text.insert(tk.END, "=== Quick Start Guide ===\n\n", "header")
            self.output_text.insert(
                tk.END, "1. Select or create an Excel file\n", "step"
            )
            self.output_text.insert(
                tk.END, "2. Select PDF papers for analysis\n", "step"
            )
            self.output_text.insert(tk.END, "3. Configure API settings\n", "step")
            self.output_text.insert(
                tk.END, '4. Click "Start Analysis" button\n', "step"
            )
        else:
            self.output_text.insert(
                tk.END, "=== 欢迎使用论文分析工具 ===\n\n", "header"
            )
            self.output_text.insert(tk.END, "=== 快速使用指南 ===\n\n", "header")
            self.output_text.insert(tk.END, "1. 选择或创建Excel文件\n", "step")
            self.output_text.insert(tk.END, "2. 选择需要分析的PDF论文文件\n", "step")
            self.output_text.insert(tk.END, "3. 配置要使用的API\n", "step")
            self.output_text.insert(tk.END, '4. 点击"开始分析"按钮\n', "step")

        self.output_text.tag_configure(
            "header", foreground="#5ba3e0", font=self.fonts["title"]
        )
        self.output_text.tag_configure("info", foreground=self.colors["fg"])
        self.output_text.tag_configure("step", foreground="#ffffff")

    def get_api_info(self):
        """获取API URL和密钥"""
        api_url = self.api_url_var.get()
        api_key = self.api_key_var.get()
        return api_url, api_key

    def load_api_configs(self):
        """加载API配置"""
        config = load_api_configs()
        self.api_url_var.set(config.get("url", ""))
        self.api_key_var.set(config.get("key", ""))
        self.api_model_var.set(config.get("model", ""))
        self.remember_key_var.set(config.get("remember", True))

    def save_api_configs(self):
        """保存API配置到文件"""
        save_api_configs(
            self.api_url_var.get(),
            self.api_key_var.get(),
            self.api_model_var.get(),
            self.remember_key_var.get(),
        )

    def _manage_excel_configs(self):
        """管理Excel表头配置"""
        from gui.excel_header_editor import show_header_editor
        from configs.excel_header_config import load_custom_columns, save_custom_columns
        import pandas as pd
        import os

        # 加载当前的自定义列
        current_columns = load_custom_columns()

        # 显示表头编辑器对话框
        result = show_header_editor(self.root, current_columns)

        if result is not None:
            # 保存到配置文件
            save_custom_columns(result)

            # 如果有选择Excel文件，则应用新表头到文件
            if self.excel_path.get() and os.path.exists(self.excel_path.get()):
                try:
                    # 尝试读取现有Excel文件
                    df = pd.read_excel(self.excel_path.get())
                    has_content = len(df) > 0  # 检查是否有数据行

                    if has_content:
                        # Excel已有内容，保留现有表头结构
                        self.status_bar["text"] = (
                            "表头配置已更新，但现有Excel表头结构未变更"
                        )
                        # 更新显示
                        self.display_excel_info(self.excel_path.get())
                    else:
                        # Excel为空，直接应用新表头
                        try:
                            # 创建空DataFrame但使用新表头
                            new_df = pd.DataFrame(columns=result)
                            # 保存回Excel
                            new_df.to_excel(self.excel_path.get(), index=False)

                            # 更新显示
                            self.display_excel_info(self.excel_path.get())
                            self.status_bar["text"] = f"表头已更新，共{len(result)}列"
                        except Exception as e:
                            print(f"更新Excel表头时出错: {str(e)}")
                            self.status_bar["text"] = (
                                f"更新Excel表头时出错: {str(e)[:50]}..."
                            )
                except Exception as e:
                    # 仅记录错误，不显示对话框
                    print(f"处理Excel文件时出错: {str(e)}")
                    self.status_bar["text"] = "表头已保存，但应用到Excel时出错"
            else:
                # 无Excel文件，仅更新状态栏
                self.status_bar["text"] = f"表头已更新，共{len(result)}列"

    def run(self):
        """运行应用程序"""
        self.root.mainloop()
