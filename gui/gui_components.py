import tkinter as tk
from tkinter import ttk


def init_styles(self):
    """初始化应用的样式"""
    try:
        ttk_module = self.ttk if hasattr(self, "ttk") else ttk

        # 创建自定义样式
        s = ttk_module.Style()

        # 按钮样式
        s.configure("TButton", font=self.fonts["button"], padding=8)
        s.configure(
            "Action.TButton",
            font=self.fonts["button"],
            padding=10,
            background=self.colors["accent"],
        )
        s.configure(
            "Cancel.TButton",
            font=self.fonts["button"],
            padding=10,
            background="#d9534f",
        )

        # 标签框架样式
        s.configure(
            "TLabelframe",
            font=self.fonts["text"],
            background=self.colors["frame"],
            padding=10,
        )
        s.configure(
            "TLabelframe.Label",
            font=self.fonts["text"],
            background=self.colors["frame"],
            foreground=self.colors["fg"],
        )

        # 标签样式
        s.configure(
            "TLabel",
            font=self.fonts["text"],
            background=self.colors["bg"],
            foreground=self.colors["fg"],
        )
        s.configure(
            "Title.TLabel",
            font=self.fonts["title"],
            background=self.colors["bg"],
            foreground=self.colors["accent"],
        )

        # 输入框样式
        s.configure(
            "TEntry",
            font=self.fonts["text"],
            padding=5,
            fieldbackground=self.colors["entry_bg"],
            foreground=self.colors["fg"],
        )

        # 复选框样式
        s.configure(
            "TCheckbutton",
            font=self.fonts["text"],
            background=self.colors["bg"],
            foreground=self.colors["fg"],
        )

        # 框架样式
        s.configure("TFrame", background=self.colors["bg"])
        s.configure("Secondary.TFrame", background=self.colors["secondary_bg"])
        s.configure("Header.TFrame", background=self.colors["secondary_bg"], padding=5)
        s.configure("Output.TFrame", background=self.colors["secondary_bg"], padding=2)

        return s
    except Exception as e:
        print(f"初始化样式时出错: {str(e)}")
        return None


def create_tools_frame(self, parent_frame):
    """创建辅助工具区域"""
    ttk_module = self.ttk if hasattr(self, "ttk") else ttk

    # 检查当前语言
    is_english = hasattr(self, "language_var") and self.language_var.get() == "English"

    tools_frame = ttk_module.LabelFrame(
        parent_frame,
        text=" Auxiliary Tools " if is_english else " 辅助工具 ",
        padding=12
    )

    # 移除论文搜索工具按钮及相关代码
    # 创建一个空的框架保持布局结构
    empty_frame = ttk_module.Frame(tools_frame)
    empty_frame.pack(fill="x", pady=5)
    
    # 添加提示标签
    info_label = ttk_module.Label(
        tools_frame, 
        text="更多工具正在开发中...",
        foreground="#888888"
    )
    info_label.pack(fill="x", pady=10)

    return tools_frame
