"""
线程安全的工具函数
"""

import threading
import tkinter as tk
import time


class ThreadSafeGUI:
    """
    提供线程安全的GUI操作工具
    用于在非GUI线程中安全地更新GUI元素
    """

    def __init__(self, root):
        self.root = root
        self.lock = threading.Lock()

    def add_task(self, func, *args, **kwargs):
        """将任务添加到GUI线程的事件循环中执行"""
        with self.lock:
            if self.root.winfo_exists():
                self.root.after(0, lambda: func(*args, **kwargs))
                return True
        return False


class ThreadSafeText:
    """
    线程安全的文本框操作工具
    用于在非GUI线程中安全地更新文本框内容
    """

    def __init__(self, text_widget, root=None):
        self.text = text_widget
        self.root = root if root else self._find_root(text_widget)
        self.lock = threading.Lock()
        self.buffer = ""  # 添加缓冲区
        self.buffer_size = 10  # 减少缓冲区大小，提高响应性
        self.last_update_time = time.time()
        self.update_interval = 0.05  # 减少更新间隔，提高流畅性

    def _find_root(self, widget):
        """找到小部件的根窗口"""
        parent = widget.master
        while parent.master:
            parent = parent.master
        return parent

    def insert(self, index, text, *tags):
        """线程安全地向文本框插入内容"""
        with self.lock:
            if self.root.winfo_exists() and self.text.winfo_exists():
                # 如果是流式输出（使用tk.END），则使用缓冲区
                if index == tk.END and not tags:
                    self.buffer += text
                    current_time = time.time()

                    # 当buffer达到一定大小或经过一定时间时，立即执行实际的insert操作
                    if (
                        len(self.buffer) >= self.buffer_size
                        or (current_time - self.last_update_time) > self.update_interval
                    ):
                        # 改为直接执行，而不是使用after
                        try:
                            if self.text.winfo_exists():
                                self.text.insert(index, self.buffer)
                                self.text.see(tk.END)
                                self.text.update()  # 强制立即更新
                                self.buffer = ""
                                self.last_update_time = current_time
                        except tk.TclError:
                            # 失败时退回到使用after方式
                            self.root.after(0, lambda: self._insert(index, self.buffer))
                            self.buffer = ""
                            self.last_update_time = current_time
                else:
                    # 正常插入，不使用缓冲
                    self.root.after(0, lambda: self._insert(index, text, *tags))

    def flush(self):
        """强制刷新缓冲区"""
        with self.lock:
            if self.buffer and self.root.winfo_exists() and self.text.winfo_exists():
                self.root.after(0, lambda: self._insert(tk.END, self.buffer))
                self.buffer = ""
                self.last_update_time = time.time()

    def _insert(self, index, text, *tags):
        """实际执行插入操作"""
        try:
            if self.text.winfo_exists():
                self.text.insert(index, text, *tags)
                self.text.see(tk.END)
                self.text.update_idletasks()
        except tk.TclError:
            pass  # 忽略可能的Tcl错误

    def see(self, index):
        """线程安全地滚动文本框"""
        with self.lock:
            if self.root.winfo_exists() and self.text.winfo_exists():
                self.root.after(0, lambda: self._see(index))

    def _see(self, index):
        """实际执行滚动操作"""
        try:
            if self.text.winfo_exists():
                self.text.see(index)
                self.text.update_idletasks()
        except tk.TclError:
            pass  # 忽略可能的Tcl错误

    def tag_configure(self, tag_name, **kwargs):
        """线程安全地配置文本标签样式"""
        with self.lock:
            if self.root.winfo_exists() and self.text.winfo_exists():
                self.root.after(0, lambda: self._tag_configure(tag_name, **kwargs))

    def _tag_configure(self, tag_name, **kwargs):
        """实际执行标签配置操作"""
        try:
            if self.text.winfo_exists():
                self.text.tag_configure(tag_name, **kwargs)
        except tk.TclError:
            pass  # 忽略可能的Tcl错误
