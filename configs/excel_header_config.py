import os
import json


def get_default_columns():
    """获取默认表头列"""
    return [
        "论文年份",
        "论文英文引用信息",
        "论文摘要（中文）",
        "研究问题及创新点",
        "研究意义",
        "研究对象及特点",
        "实验设置",
        "数据分析方法",
        "重要结论",
        "未来研究展望",
    ]


def get_config_path():
    """获取配置文件路径"""
    user_docs = os.path.join(os.path.expanduser("~"), "Documents", "论文分析工具")
    os.makedirs(user_docs, exist_ok=True)
    return os.path.join(user_docs, "excel_columns.json")


def load_custom_columns():
    """加载自定义表头"""
    config_path = get_config_path()

    try:
        if os.path.exists(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                columns = data.get("columns", [])
                if columns:
                    return columns
    except Exception as e:
        print(f"读取自定义表头出错: {e}")

    # 如果加载失败或没有自定义表头，返回默认表头
    return get_default_columns()


def save_custom_columns(columns):
    """保存自定义表头"""
    config_path = get_config_path()

    # 确保目录存在
    os.makedirs(os.path.dirname(config_path), exist_ok=True)

    try:
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump({"columns": columns}, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"保存自定义表头出错: {e}")
        return False
