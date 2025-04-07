import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from utils.canvas_utils import ScrollableFrame

# 修复导入错误：正确导入create_actions_frame和create_tools_frame
from gui.gui_actions import create_actions_frame
from gui.gui_components import create_tools_frame


def setup_ui(self):
    ttk_module = self.ttk if hasattr(self, "ttk") else ttk

    # 创建主分隔窗口，明确设置最小尺寸
    main_paned = ttk_module.PanedWindow(self.root, orient=tk.HORIZONTAL)
    main_paned.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

    # 左侧容器设置固定宽度并防止自动缩放
    left_container = ttk_module.Frame(main_paned, width=400, relief="flat")
    left_container.pack_propagate(False)  # 防止子组件影响容器大小
    main_paned.add(left_container, weight=0)  # 设置weight为0使其保持固定大小

    try:
        scrollable_frame = ScrollableFrame(left_container, background=self.colors["bg"])
        scrollable_frame.pack(fill=tk.BOTH, expand=True)
        left_frame = scrollable_frame.frame
    except Exception:
        left_frame = ttk_module.Frame(left_container)
        left_frame.pack(fill=tk.BOTH, expand=True)

    # 添加语言切换按钮
    lang_frame = ttk_module.Frame(left_frame)
    lang_frame.pack(fill="x", pady=(0, 15), padx=5)

    # 确保语言按钮文本与当前语言一致
    is_english = self.language_var.get() == "English"
    lang_btn_text = "切换为中文" if is_english else "Switch to English"

    lang_btn = ttk_module.Button(
        lang_frame,
        text=lang_btn_text,
        command=self.toggle_language,
        style="Action.TButton",
    )
    lang_btn.pack(fill="x")
    self.lang_btn = lang_btn  # 保存引用以便后续更新

    # 更新Excel框架标题根据当前语言
    excel_frame_title = " Excel File Management " if is_english else " Excel文件管理 "

    excel_frame = ttk_module.LabelFrame(left_frame, text=excel_frame_title, padding=12)
    excel_frame.pack(fill="x", pady=(0, 15), padx=5)

    excel_btn_frame = ttk_module.Frame(excel_frame)
    excel_btn_frame.pack(fill="x", pady=(0, 10))

    # 所有按钮文本使用当前语言
    select_excel_btn_text = "Select Excel File" if is_english else "选择Excel文件"
    create_excel_btn_text = "Create Excel File" if is_english else "新建Excel文件"

    select_excel_btn = ttk_module.Button(
        excel_btn_frame, text=select_excel_btn_text, command=self.select_excel
    )
    select_excel_btn.pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 5))

    create_excel_btn = ttk_module.Button(
        excel_btn_frame, text=create_excel_btn_text, command=self.create_excel
    )
    create_excel_btn.pack(side=tk.RIGHT, fill="x", expand=True)

    # 添加自定义表头按钮
    custom_header_btn = ttk_module.Button(
        excel_frame, text="自定义表头", command=self.manage_excel_configs
    )
    custom_header_btn.pack(fill="x", pady=(5, 0))

    self.excel_info_label = ttk_module.Label(
        excel_frame,
        text="尚未选择Excel文件",
        wraplength=350,
        justify="left",
        foreground="#c0c0c0",
    )
    self.excel_info_label.pack(fill="x", pady=(5, 0))

    pdf_frame = ttk_module.LabelFrame(left_frame, text=" PDF文件处理 ", padding=12)
    pdf_frame.pack(fill="x", pady=(0, 15), padx=5)

    # 移除原来的"选择PDF文件"按钮
    # 直接创建PDF文件管理按钮
    pdf_manage_frame = ttk_module.Frame(pdf_frame)
    pdf_manage_frame.pack(fill="x", pady=(0, 0))

    add_pdf_btn = ttk_module.Button(
        pdf_manage_frame, text="添加PDF文件", command=self.add_pdf
    )
    add_pdf_btn.pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 2))

    remove_pdf_btn = ttk_module.Button(
        pdf_manage_frame, text="删除选中PDF", command=self.remove_pdf
    )
    remove_pdf_btn.pack(side=tk.RIGHT, fill="x", expand=True, padx=(2, 0))

    api_frame = ttk_module.LabelFrame(left_frame, text=" API设置 ", padding=12)
    api_frame.pack(fill="x", pady=(0, 15), padx=5)

    # 删除API供应商下拉框，改为直接输入模型名称
    model_label = ttk_module.Label(api_frame, text="模型名称:")
    model_label.pack(anchor=tk.W, pady=(0, 5))
    self.api_model_entry = ttk_module.Entry(
        api_frame, textvariable=self.api_model_var, font=self.fonts["text"]
    )
    self.api_model_entry.pack(fill="x", pady=(0, 10))

    # API URL
    api_url_label = ttk_module.Label(api_frame, text="API URL:")
    api_url_label.pack(anchor=tk.W, pady=(0, 5))
    self.api_url_entry = ttk_module.Entry(
        api_frame, textvariable=self.api_url_var, font=self.fonts["text"]
    )
    self.api_url_entry.pack(fill="x", pady=(0, 10))

    # API Key
    api_key_label = ttk_module.Label(api_frame, text="API Key:")
    api_key_label.pack(anchor=tk.W, pady=(0, 5))
    self.api_key_entry = ttk_module.Entry(
        api_frame, textvariable=self.api_key_var, font=self.fonts["text"], show="*"
    )
    self.api_key_entry.pack(fill="x", pady=(0, 5))

    # 记住API设置复选框
    self.remember_key_check = ttk_module.Checkbutton(
        api_frame,
        text="记住API设置",
        variable=self.remember_key_var,
        style="TCheckbutton",
    )
    self.remember_key_check.pack(anchor=tk.W, pady=(5, 0))

    # 添加操作按钮区域
    actions_frame = self.create_actions_frame(left_frame)
    actions_frame.pack(fill="x", pady=(0, 15), padx=5)

    # 添加工具区域
    tools_frame = self.create_tools_frame(left_frame)
    tools_frame.pack(fill="x", pady=(0, 5), padx=5)

    # 添加状态栏
    status_frame = ttk_module.Frame(self.root, padding=5)
    status_frame.pack(side=tk.BOTTOM, fill=tk.X)

    self.status_bar = ttk_module.Label(
        status_frame, text="就绪", anchor=tk.W, padding=(10, 5)
    )
    self.status_bar.pack(fill=tk.X)

    # 右侧框架使用更明显的视觉边界
    right_container = ttk_module.Frame(main_paned)
    main_paned.add(right_container, weight=1)  # 设置weight为1使其获得所有可用空间

    # 添加右侧标题栏，使其更加明显
    right_header = ttk_module.Frame(right_container, style="Header.TFrame")
    right_header.pack(fill=tk.X, pady=(0, 10))

    output_label = ttk_module.Label(
        right_header,
        text="分析输出",
        anchor=tk.W,
        font=self.fonts["title"],
        foreground=self.colors["accent"],
    )
    output_label.pack(fill=tk.X, pady=(5, 5), padx=5)

    # 右侧内容区域
    right_frame = ttk_module.Frame(right_container, style="Secondary.TFrame")
    right_frame.pack(fill=tk.BOTH, expand=True)

    # 使用带边框的容器包装文本区域，增强视觉边界
    output_frame = ttk_module.Frame(
        right_frame, style="Output.TFrame", borderwidth=1, relief="groove"
    )
    output_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    text_scroll = ttk_module.Scrollbar(output_frame)
    text_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    self.output_text = tk.Text(
        output_frame,
        yscrollcommand=text_scroll.set,
        wrap=tk.WORD,
        bg=self.colors["secondary_bg"],
        fg=self.colors["fg"],
        font=self.fonts["console"],
        padx=10,
        pady=10,  # 添加内部填充，改善文本可读性
    )
    self.output_text.pack(fill=tk.BOTH, expand=True)

    text_scroll.config(command=self.output_text.yview)

    self.output_text.tag_configure(
        "header", foreground="#4a8cca", font=self.fonts["title"]
    )
    self.output_text.tag_configure("info", foreground=self.colors["fg"])

    # 在root更新后强制设置分隔条位置
    self.root.update_idletasks()

    # 使用更可靠的方式设置分隔条位置
    def set_sash_position():
        width = main_paned.winfo_width()
        main_paned.sashpos(0, int(width * 0.25))

    self.root.after(200, set_sash_position)

    return {
        "main_paned": main_paned,
        "left_frame": left_frame,
        "right_frame": right_frame,
        "output_text": self.output_text,
    }
