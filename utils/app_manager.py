"""
管理应用程序全局服务和资源
"""

import threading

# 存储应用程序的全局服务和资源
_app_services = {}
_app_lock = threading.Lock()

# 统一的任务取消标志
_analysis_cancelled = False
_terminate_all_tasks = False  # 新增统一的任务终止标志

def init_app_services(root):
    """初始化应用程序服务"""
    global _app_services, _analysis_cancelled, _terminate_all_tasks
    with _app_lock:
        _app_services["root"] = root
        _app_services["threads"] = []
        _analysis_cancelled = False
        _terminate_all_tasks = False


def cleanup_app_services():
    """清理应用程序服务"""
    global _app_services, _analysis_cancelled, _terminate_all_tasks
    with _app_lock:
        # 终止所有线程
        terminate_specific_threads()
        _app_services.clear()
        _analysis_cancelled = False
        _terminate_all_tasks = False


def terminate_specific_threads():
    """终止特定线程"""
    # 这里可以添加终止特定线程的逻辑
    pass


def get_thread_safe_gui(root):
    """获取线程安全的GUI交互对象"""
    from utils.thread_utils import ThreadSafeGUI
    return ThreadSafeGUI(root)


def set_analysis_cancelled(flag):
    """设置分析取消标志"""
    global _analysis_cancelled
    _analysis_cancelled = flag
    # 同时设置终止标志，确保一致性
    if flag:
        global _terminate_all_tasks
        _terminate_all_tasks = flag


def is_analysis_cancelled():
    """检查分析是否被取消"""
    global _analysis_cancelled
    return _analysis_cancelled


def set_terminate_all_tasks(flag):
    """设置终止所有任务的标志"""
    global _terminate_all_tasks
    _terminate_all_tasks = flag
    # 同时设置分析取消标志，确保一致性
    if flag:
        global _analysis_cancelled
        _analysis_cancelled = flag


def is_terminate_all_tasks():
    """检查是否应该终止所有任务"""
    global _terminate_all_tasks
    return _terminate_all_tasks
