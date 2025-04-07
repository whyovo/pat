import os
import tkinter as tk
from tkinter import ttk, messagebox
from configs.excel_header_config import get_default_columns


def show_header_editor(parent, current_columns=None):
    """显示Excel表头编辑器对话框"""
    # 如果没有提供当前列，使用默认列
    if current_columns is None:
        current_columns = get_default_columns()

    # 创建对话框
    dialog = tk.Toplevel(parent)
    dialog.title("Excel表头编辑器")
    dialog.transient(parent)  # 设置为父窗口的临时窗口
    dialog.grab_set()  # 模态对话框

    # 居中显示
    parent_x = parent.winfo_rootx()
    parent_y = parent.winfo_rooty()
    parent_width = parent.winfo_width()
    parent_height = parent.winfo_height()

    dialog_width = 500
    dialog_height = 600

    x = parent_x + (parent_width - dialog_width) // 2
    y = parent_y + (parent_height - dialog_height) // 2

    dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")

    # 结果变量
    result_var = [None]

    # 主框架
    main_frame = ttk.Frame(dialog, padding=20)
    main_frame.pack(fill=tk.BOTH, expand=True)

    # 标题
    title_label = ttk.Label(
        main_frame, text="编辑Excel表头字段", font=("微软雅黑", 14, "bold")
    )
    title_label.pack(pady=(0, 10))

    # 说明
    desc_label = ttk.Label(
        main_frame,
        text="您可以添加、删除或重新排序Excel表头字段。\n这些字段将用于保存论文分析结果。",
        justify=tk.CENTER,
    )
    desc_label.pack(pady=(0, 5))

    # 添加表头来源提示
    default_columns = get_default_columns()
    is_default = sorted(current_columns) == sorted(default_columns)

    status_text = "当前使用: " + (
        "默认表头" if is_default else "自定义表头 (来自Excel)"
    )
    status_label = ttk.Label(
        main_frame,
        text=status_text,
        foreground="green" if is_default else "orange",
        font=("微软雅黑", 9, "italic"),
        justify=tk.CENTER,
    )
    status_label.pack(pady=(0, 15))

    # 列表框架
    list_frame = ttk.Frame(main_frame)
    list_frame.pack(fill=tk.BOTH, expand=True)

    # 创建列表和滚动条
    scrollbar = ttk.Scrollbar(list_frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    columns_listbox = tk.Listbox(
        list_frame, selectmode=tk.SINGLE, height=15, font=("微软雅黑", 11)
    )
    columns_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    columns_listbox.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=columns_listbox.yview)

    # 填充列表
    for col in current_columns:
        columns_listbox.insert(tk.END, col)

    # 按钮框架
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill=tk.X, pady=(10, 0))

    # 创建输入框和添加按钮
    input_frame = ttk.Frame(button_frame)
    input_frame.pack(fill=tk.X, pady=5)

    ttk.Label(input_frame, text="新字段名:").pack(side=tk.LEFT)

    new_column_var = tk.StringVar()
    new_column_entry = ttk.Entry(input_frame, textvariable=new_column_var, width=30)
    new_column_entry.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

    def add_column():
        column_name = new_column_var.get().strip()
        if not column_name:
            messagebox.showwarning("警告", "请输入字段名")
            return

        if column_name in columns_listbox.get(0, tk.END):
            messagebox.showwarning("警告", f"字段 '{column_name}' 已存在")
            return

        columns_listbox.insert(tk.END, column_name)
        new_column_var.set("")  # 清空输入框

    add_button = ttk.Button(input_frame, text="添加", command=add_column)
    add_button.pack(side=tk.LEFT)

    # 移动和删除按钮框架
    move_frame = ttk.Frame(button_frame)
    move_frame.pack(fill=tk.X, pady=5)

    def move_up():
        selected = columns_listbox.curselection()
        if not selected or selected[0] == 0:
            return

        idx = selected[0]
        value = columns_listbox.get(idx)

        columns_listbox.delete(idx)
        columns_listbox.insert(idx - 1, value)
        columns_listbox.selection_set(idx - 1)

    def move_down():
        selected = columns_listbox.curselection()
        if not selected or selected[0] == columns_listbox.size() - 1:
            return

        idx = selected[0]
        value = columns_listbox.get(idx)

        columns_listbox.delete(idx)
        columns_listbox.insert(idx + 1, value)
        columns_listbox.selection_set(idx + 1)

    def delete_item():
        selected = columns_listbox.curselection()
        if not selected:
            return

        idx = selected[0]
        columns_listbox.delete(idx)

        if idx < columns_listbox.size():
            columns_listbox.selection_set(idx)
        elif columns_listbox.size() > 0:
            columns_listbox.selection_set(columns_listbox.size() - 1)

    def reset_to_default():
        if messagebox.askyesno(
            "确认", "确定要恢复默认表头吗？所有自定义字段将被替换。"
        ):
            columns_listbox.delete(0, tk.END)
            default = get_default_columns()
            for col in default:
                columns_listbox.insert(tk.END, col)

    ttk.Button(move_frame, text="上移", command=move_up).pack(side=tk.LEFT, padx=2)
    ttk.Button(move_frame, text="下移", command=move_down).pack(side=tk.LEFT, padx=2)
    ttk.Button(move_frame, text="删除", command=delete_item).pack(side=tk.LEFT, padx=2)
    ttk.Button(move_frame, text="恢复默认", command=lambda: reset_to_default()).pack(
        side=tk.RIGHT, padx=2
    )

    # 确认和取消按钮框架
    action_frame = ttk.Frame(main_frame)
    action_frame.pack(fill=tk.X, pady=(20, 0))

    def on_save():
        columns = list(columns_listbox.get(0, tk.END))
        if not columns:
            messagebox.showerror("错误", "表头列表不能为空")
            return
        result_var[0] = columns
        dialog.destroy()

    ttk.Button(action_frame, text="确认", command=on_save, padding=10).pack(
        side=tk.LEFT, padx=5
    )
    ttk.Button(action_frame, text="取消", command=dialog.destroy, padding=10).pack(
        side=tk.RIGHT, padx=5
    )

    # 等待对话框关闭
    parent.wait_window(dialog)

    return result_var[0]
