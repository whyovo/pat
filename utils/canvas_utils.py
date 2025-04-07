import tkinter as tk
from tkinter import ttk


class ScrollableFrame:
    """
    创建一个可滚动的框架，适用于内容过多的情况
    """

    def __init__(self, container, background="#2d2d2d"):
        # 创建主容器框架，用于容纳画布和滚动条
        self.main_frame = ttk.Frame(container)

        # 创建画布和滚动条
        self.canvas = tk.Canvas(
            self.main_frame, background=background, highlightthickness=0
        )
        self.scrollbar = ttk.Scrollbar(
            self.main_frame, orient="vertical", command=self.canvas.yview
        )

        # 配置画布
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # 创建实际内容框架
        self.frame = ttk.Frame(self.canvas)
        self.frame_id = self.canvas.create_window(
            (0, 0), window=self.frame, anchor="nw"
        )

        # 布局组件 - 默认不显示滚动条
        self.scrollbar.pack_forget()  # 初始时不显示滚动条
        self.canvas.pack(side="left", fill="both", expand=True)

        # 绑定事件
        self.frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # 绑定鼠标滚轮滚动
        self._bind_mousewheel()

        # 存储容器引用，用于检查大小
        self.container = container

    def pack(self, **kwargs):
        """实现pack方法，代理到主框架"""
        self.main_frame.pack(**kwargs)
        return self

    def grid(self, **kwargs):
        """实现grid方法，代理到主框架"""
        self.main_frame.grid(**kwargs)
        return self

    def place(self, **kwargs):
        """实现place方法，代理到主框架"""
        self.main_frame.place(**kwargs)
        return self

    def pack_forget(self):
        """实现pack_forget方法"""
        self.main_frame.pack_forget()
        return self

    def grid_forget(self):
        """实现grid_forget方法"""
        self.main_frame.grid_forget()
        return self

    def place_forget(self):
        """实现place_forget方法"""
        self.main_frame.place_forget()
        return self

    def _on_frame_configure(self, event):
        """更新滚动区域以匹配框架的大小并检查是否需要滚动条"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.update_scrollbar()

    def _on_canvas_configure(self, event):
        """调整内部框架宽度以匹配画布宽度并检查是否需要滚动条"""
        # 调整框架宽度以匹配画布宽度
        canvas_width = event.width
        self.canvas.itemconfig(self.frame_id, width=canvas_width)

        # 检查是否需要滚动条
        self.update_scrollbar()

    def update_scrollbar(self):
        """根据内容高度决定是否显示滚动条"""
        # 获取画布高度和内容框架高度
        canvas_height = self.canvas.winfo_height()
        frame_height = self.frame.winfo_reqheight()

        # 如果内容高度大于画布高度，显示滚动条
        if frame_height > canvas_height:
            if not self.scrollbar.winfo_ismapped():
                self.scrollbar.pack(side="right", fill="y")
                # 当添加滚动条时，重新设置画布宽度
                self.canvas.pack_forget()
                self.canvas.pack(side="left", fill="both", expand=True)
        else:
            # 如果内容高度小于等于画布高度，隐藏滚动条
            if self.scrollbar.winfo_ismapped():
                self.scrollbar.pack_forget()
                # 当移除滚动条时，重新设置画布宽度
                self.canvas.pack_forget()
                self.canvas.pack(side="left", fill="both", expand=True)

    def _bind_mousewheel(self):
        """绑定鼠标滚轮事件"""
        # Windows
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel_windows)
        # Linux
        self.canvas.bind_all("<Button-4>", self._on_mousewheel_linux)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel_linux)

    def _on_mousewheel_windows(self, event):
        """Windows鼠标滚轮滚动，仅在有滚动条时响应"""
        if self.canvas.winfo_exists() and self.scrollbar.winfo_ismapped():
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_mousewheel_linux(self, event):
        """Linux鼠标滚轮滚动，仅在有滚动条时响应"""
        if self.canvas.winfo_exists() and self.scrollbar.winfo_ismapped():
            if event.num == 4:
                self.canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.canvas.yview_scroll(1, "units")

    def unbind_mousewheel(self):
        """解绑鼠标滚轮事件，在框架销毁前调用"""
        try:
            self.canvas.unbind_all("<MouseWheel>")
            self.canvas.unbind_all("<Button-4>")
            self.canvas.unbind_all("<Button-5>")
        except:
            pass
