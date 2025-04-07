import tkinter as tk
import pandas as pd
import os
import sys

# 按钮文本映射 - 添加更多按钮文本的映射关系
BUTTON_TEXT_MAPPING = {
    # 操作按钮
    "开始分析": "Start Analysis",
    "重新生成": "Regenerate",
    "创新点评估": "Innovation Assessment",
    "综述生成": "Review Generation",
    "取消任务": "Cancel Task",
    # 文件操作按钮
    "选择Excel文件": "Select Excel File",
    "新建Excel文件": "Create Excel File",
    "添加PDF文件": "Add PDF Files",
    "删除选中PDF": "Delete Selected PDF",
    "自定义表头": "Custom Headers",
    # 新增 Excel 和 PDF 操作相关按钮文本
    "选择PDF文件": "Select PDF Files",
    "管理PDF文件": "Manage PDF Files",
    "排序": "Sort",
    "筛选": "Filter",
    "导出": "Export",
    "保存设置": "Save Settings",
    "取消": "Cancel",
    "确认": "Confirm",
    "继续": "Continue",
    # 标签框架标题
    " Excel文件管理 ": " Excel File Management ",
    " PDF文件处理 ": " PDF File Processing ",
    " API设置 ": " API Settings ",
    " 论文分析操作 ": " Paper Analysis Operations ",
    " 辅助工具 ": " Auxiliary Tools ",
    # 标签和提示
    "尚未选择Excel文件": "No Excel file selected",
    "请输入API信息": "Please enter API information",
    "模型名称:": "Model Name:",
    "API URL:": "API URL:",
    "API Key:": "API Key:",
    "记住API设置": "Remember API Settings",
    # 语言切换按钮
    "Switch to English": "切换为中文",
    "切换为中文": "Switch to English",
}

# 英文到中文的反向映射
REVERSE_BUTTON_TEXT_MAPPING = {v: k for k, v in BUTTON_TEXT_MAPPING.items()}


def disable_analysis_buttons(self):
    """禁用分析按钮"""

    def disable_button(button):
        if hasattr(self, button):
            getattr(self, button)["state"] = "disabled"

    disable_button("start_analysis_btn")
    disable_button("regenerate_btn")
    disable_button("extract_btn")
    disable_button("review_btn")


def enable_analysis_buttons(self, keep_regenerate_disabled=False):
    """启用分析按钮"""

    def enable_button(button):
        if hasattr(self, button):
            getattr(self, button)["state"] = "normal"

    enable_button("start_analysis_btn")
    if not keep_regenerate_disabled:
        enable_button("regenerate_btn")
    enable_button("extract_btn")
    enable_button("review_btn")


def _process_buttons(self, callback):
    """处理所有按钮的通用函数"""
    for widget_name in dir(self):
        if widget_name.endswith(("_btn", "_button")) and hasattr(self, widget_name):
            widget = getattr(self, widget_name)
            if hasattr(widget, "cget") and widget.winfo_exists():
                try:
                    callback(widget)
                except Exception as e:
                    print(f"处理按钮 {widget_name} 时出错: {str(e)}")


def update_button_texts(self):
    """根据当前语言更新按钮文本"""
    is_english = self.language_var.get() == "English"

    def update_text(btn):
        try:
            current_text = btn.cget("text")
            if is_english and current_text in BUTTON_TEXT_MAPPING:
                btn.configure(text=BUTTON_TEXT_MAPPING[current_text])
            elif not is_english and current_text in REVERSE_BUTTON_TEXT_MAPPING:
                btn.configure(text=REVERSE_BUTTON_TEXT_MAPPING[current_text])
        except:
            pass

    # 递归处理所有子组件，而不只是成员变量
    def process_all_widgets(parent):
        # 遍历所有子组件
        for child in parent.winfo_children():
            # 如果是按钮类型，更新其文本
            if child.winfo_class() in ("Button", "TButton"):
                update_text(child)
            # 如果是标签框架，更新其文本
            elif child.winfo_class() == "TLabelframe":
                label_text = child.cget("text")
                if is_english and label_text in BUTTON_TEXT_MAPPING:
                    child.configure(text=BUTTON_TEXT_MAPPING[label_text])
                elif not is_english and label_text in REVERSE_BUTTON_TEXT_MAPPING:
                    child.configure(text=REVERSE_BUTTON_TEXT_MAPPING[label_text])
            # 递归处理子组件
            process_all_widgets(child)

    # 首先处理已知的类成员按钮 - 保持原有逻辑以确保向后兼容
    _process_buttons(self, update_text)
    
    # 然后递归处理整个UI树中的所有按钮
    process_all_widgets(self.root)

    # 更新特定的Label文本 - 如Excel信息标签等
    if hasattr(self, "excel_info_label") and self.excel_info_label.winfo_exists():
        if self.excel_path.get():
            # 保留路径信息，只翻译标签部分
            path = self.excel_path.get()
            filename = os.path.basename(path)
            if is_english:
                self.excel_info_label.config(
                    text=f"Current file: {filename}\nPath: {path}"
                )
            else:
                self.excel_info_label.config(text=f"当前文件: {filename}\n路径: {path}")
        else:
            # 没有选择文件时
            text = "No Excel file selected" if is_english else "尚未选择Excel文件"
            self.excel_info_label.config(text=text)


def update_progress_status(self, current_index, total_files):
    """
    更新状态栏
    """
    if hasattr(self, "status_bar"):
        percentage = int((current_index / total_files) * 100)
        self.status_bar["text"] = (
            f"进度: {percentage}% - 处理第 {current_index}/{total_files} 个文件"
        )


def append_response_chunk(self, text_chunk):
    """
    将API响应块添加到输出文本，确保实时显示
    """
    if hasattr(self, "output_text") and self.output_text.winfo_exists():
        try:
            # 直接在主线程中执行更新，避免延迟
            self.output_text.insert(tk.END, text_chunk, "result")
            self.output_text.see(tk.END)
            # 这里是关键：强制更新界面显示
            self.output_text.update()  # 强制立即更新显示
        except tk.TclError as e:
            print(f"更新文本时出现TclError: {str(e)}")
            # 通过after机制安全地添加文本
            try:
                self.root.after(0, lambda: self._safe_append_chunk(text_chunk))
            except Exception as e2:
                print(f"使用after机制添加文本时出错: {str(e2)}")


def _safe_append_chunk(self, text_chunk):
    """安全地添加文本块的备用方法"""
    try:
        if self.output_text.winfo_exists():
            self.output_text.insert(tk.END, text_chunk, "result")
            self.output_text.see(tk.END)
            self.output_text.update()  # 强制立即更新
    except Exception as e:
        print(f"安全添加文本块时出错: {str(e)}")


# 添加一个全局标志，确保save_excel_result函数只被调用一次
_excel_save_in_progress = False


def save_excel_result(self, df):
    """保存Excel结果，确保只执行一次"""
    global _excel_save_in_progress

    # 检查是否已经在保存中，避免重复保存
    if _excel_save_in_progress:
        print("Excel保存已在进行中，跳过重复调用")
        return True

    # 标记为保存中
    _excel_save_in_progress = True

    try:
        # 取消状态检查
        if (
            hasattr(self, "cancel_analysis_requested")
            and self.cancel_analysis_requested
        ):
            print("检测到取消状态，不保存Excel结果")
            self.output_text.insert(
                tk.END, "\n已取消分析，不保存Excel结果。\n", "warning"
            )
            self.output_text.tag_configure(
                "warning", foreground="#e69138", font=("微软雅黑", 10, "bold")
            )
            self.output_text.see(tk.END)
            return False
        try:
            # 导入必要模块
            import sys, os

            current_dir = os.path.dirname(os.path.abspath(__file__))
            if current_dir not in sys.path:
                sys.path.append(current_dir)
            from configs.excel_header_config import load_custom_columns
            from utils.excel_utils import save_to_excel_with_format, format_excel_file

            # 确保包含所有所需列
            try:
                current_columns = load_custom_columns()
                for col in current_columns:
                    if (col not in df.columns) and (
                        col != "创新点评估" and col != "综述生成"
                    ):
                        df[col] = "未提供相关信息"
            except Exception as e:
                print(f"加载表头配置时出错: {e}")

            # 读取并合并数据
            excel_path = self.excel_path.get()
            if os.path.exists(excel_path):
                try:
                    existing_df = pd.read_excel(excel_path)
                    combined_df = pd.concat([existing_df, df], ignore_index=True)
                except Exception as e:
                    print(f"读取Excel文件出错: {e}")
                    combined_df = df.copy()
            else:
                combined_df = df.copy()

            # 按年份排序
            try:
                if ("论文年份" in combined_df.columns):
                    # 确保年份是数值类型
                    combined_df["论文年份"] = pd.to_numeric(
                        combined_df["论文年份"], errors="coerce"
                    )
                    # 按年份降序排序（最新的在前）
                    combined_df = combined_df.sort_values(
                        by="论文年份", ascending=False
                    ).reset_index(drop=True)
                    print("已按论文年份排序")
                    # 添加排序成功的提示
                    self.output_text.insert(
                        tk.END, "\n已按论文年份降序排列（最新的在前）\n", "info"
                    )
                    self.output_text.tag_configure("info", foreground="#4a8cca")
            except Exception as e:
                print(f"排序数据时出错: {e}")

            # 保存数据
            save_success = save_to_excel_with_format(
                combined_df, excel_path, append_mode=False
            )

            # 无论保存模式如何，都尝试格式化Excel
            if save_success:
                self.output_text.insert(tk.END, "\n正在应用Excel格式美化...\n")
                self.output_text.see(tk.END)
                self.output_text.update()

                try:
                    # 确保Excel文件存在
                    if os.path.exists(excel_path):
                        # 强制执行格式化
                        format_result = format_excel_file(excel_path)
                        if format_result:
                            self.output_text.insert(
                                tk.END,
                                "\n✓ 已保存分析结果到Excel并完成格式美化\n",
                                "success",
                            )
                        else:
                            self.output_text.insert(
                                tk.END,
                                "\n✓ 已保存分析结果到Excel，但格式美化可能不完整\n",
                                "warning",
                            )

                        self.output_text.tag_configure(
                            "success",
                            foreground="#4caf50",
                            font=("微软雅黑", 10, "bold"),
                        )
                        self.analysis_completed = True
                    else:
                        self.output_text.insert(
                            tk.END,
                            f"\n保存成功，但无法找到Excel文件进行格式化：{excel_path}\n",
                            "warning",
                        )
                except Exception as format_error:
                    print(f"Excel格式化出错: {str(format_error)}")
                    import traceback

                    print(traceback.format_exc())
                    self.output_text.insert(
                        tk.END,
                        f"\n保存成功，但格式化失败：{str(format_error)}\n",
                        "warning",
                    )
                    self.output_text.tag_configure("warning", foreground="#e69138")
            else:
                self.output_text.insert(
                    tk.END,
                    f"\n保存Excel时可能发生问题，请检查文件权限和格式。\n",
                    "warning",
                )

            self.output_text.see(tk.END)
            return save_success
        except Exception as e:
            print(f"保存Excel错误详情: {str(e)}")
            import traceback

            print(traceback.format_exc())
            self.output_text.insert(tk.END, f"\n保存Excel时出错：{str(e)}\n", "error")
            self.output_text.tag_configure("error", foreground="#cc0000")
            return False
    finally:
        # 完成后重置标志
        _excel_save_in_progress = False


def on_analysis_complete(self):
    """分析完成后的处理"""
    enable_analysis_buttons(self)

    # 重置取消和终止标志
    self.cancel_analysis_requested = False
    if hasattr(self, "terminate_all_tasks"):
        self.terminate_all_tasks = False
        
    # 同时重置全局标志
    from utils.app_manager import set_analysis_cancelled, set_terminate_all_tasks
    set_analysis_cancelled(False)
    set_terminate_all_tasks(False)

    # 如果存在"取消任务"按钮，把它改回"取消任务"文本并禁用它
    if hasattr(self, "cancel_analysis_btn"):
        self.cancel_analysis_btn.configure(text="取消任务", state="disabled")

    self.analysis_completed = True


def show_error_and_reset(self, error_message):
    """显示错误并重置界面状态"""
    if hasattr(self, "output_text") and self.output_text.winfo_exists():
        self.output_text.insert(tk.END, f"\n错误: {error_message}\n", "error")
        self.output_text.tag_configure("error", foreground="#ef5350")
        self.output_text.see(tk.END)

    if hasattr(self, "status_bar"):
        self.status_bar["text"] = (
            f"错误: {error_message[:50]}..."
            if len(error_message) > 50
            else f"错误: {error_message}"
        )

    enable_analysis_buttons(self)


def cancel_analysis(self):
    """取消当前分析任务"""
    if hasattr(self, "cancel_analysis_btn"):
        self.cancel_analysis_btn.configure(text="正在取消...", state="disabled")

    # 设置两种取消标志，确保一致性
    self.cancel_analysis_requested = True
    if not hasattr(self, "terminate_all_tasks"):
        self.terminate_all_tasks = True
    else:
        self.terminate_all_tasks = True
        
    self.status_bar["text"] = "正在取消任务，请等待..."

    # 导入设置取消标志的函数
    from utils.app_manager import set_analysis_cancelled, set_terminate_all_tasks

    # 同时设置两个标志
    set_analysis_cancelled(True)
    set_terminate_all_tasks(True)

    # 使用线程安全的方式更新输出
    if hasattr(self, "output_text"):
        self.output_text.insert(
            tk.END, "\n\n取消请求已发送，等待当前操作完成...\n", "warning"
        )
        self.output_text.tag_configure("warning", foreground="#e69138")
        self.output_text.see(tk.END)
