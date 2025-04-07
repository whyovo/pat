"""策略与功能实现模块"""

# 导出创新点提取和综述生成模块
from .extract import extract_content, create_analysis_report
from .review import generate_review, create_review_document

__all__ = [
    "extract_content",
    "create_analysis_report",
    "generate_review",
    "create_review_document",
]
