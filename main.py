import os
import sys
import tkinter as tk
import ttkbootstrap as ttk_bootstrap

# 添加当前路径到系统路径
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

from gui.app import PaperAnalyzer

if __name__ == "__main__":
    try:
        app = PaperAnalyzer()
        app.run()
    except Exception as e:
        import traceback
        print(f"程序启动失败: {str(e)}")
        print(traceback.format_exc())
        try:
            # 尝试显示错误窗口
            import tkinter.messagebox as msgbox
            msgbox.showerror("启动错误", f"程序启动失败:\n{str(e)}")
        except:
            pass
        input("按Enter键退出...")
