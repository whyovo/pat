import tkinter as tk
from tkinter import ttk


def create_actions_frame(self, parent_frame):
    """创建操作按钮界面"""
    # 使用传递的 ttk 模块
    ttk_module = self.ttk if hasattr(self, "ttk") else ttk

    # 检查当前语言
    is_english = hasattr(self, "language_var") and self.language_var.get() == "English"

    # 操作按钮框架 - 删除pack调用，由gui.py统一管理布局
    action_frame = ttk_module.LabelFrame(
        parent_frame,
        text=" 论文分析操作 " if not is_english else " Paper Analysis Operations ",
        padding=15,
    )

    # 主操作按钮行
    main_action_row = ttk_module.Frame(action_frame)
    main_action_row.pack(fill="x", pady=(0, 10))

    # 开始分析按钮 - 确保强调样式和明显的颜色
    self.start_analysis_btn = ttk_module.Button(
        main_action_row,
        text="开始分析",
        command=self.analyze_papers,
        style="Action.TButton",
        padding=12,
    )
    self.start_analysis_btn.pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 5))

    # 取消生成按钮 - 初始时禁用
    self.cancel_analysis_btn = ttk_module.Button(
        main_action_row,
        text="取消任务",
        command=self.cancel_analysis,
        style="Cancel.TButton",
        padding=12,
        state="disabled",
    )
    self.cancel_analysis_btn.pack(side=tk.RIGHT, fill="x", expand=True)

    # 重新生成按钮，单独一行
    self.regenerate_btn = ttk_module.Button(
        action_frame,
        text="重新生成",
        command=self.regenerate,
        style="Action.TButton",
        padding=12,
    )
    self.regenerate_btn.pack(fill="x", pady=(0, 10))

    # 创新点评估按钮 - 单独一行
    self.extract_btn = ttk_module.Button(
        action_frame,
        text="创新点评估",
        command=self.extract_content,
        style="Action.TButton",
        padding=12,
    )
    self.extract_btn.pack(fill="x", pady=(0, 10))

    # 综述生成按钮 - 单独一行
    self.review_btn = ttk_module.Button(
        action_frame,
        text="综述生成",
        command=self.generate_review,
        style="Action.TButton",
        padding=12,
    )
    self.review_btn.pack(fill="x")

    # 确保按钮在创建后可见 - 添加此行解决可能的显示问题
    action_frame.update_idletasks()

    return action_frame
