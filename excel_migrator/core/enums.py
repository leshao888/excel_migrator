"""枚举定义"""
from enum import Enum


class Direction(str, Enum):
    """数据复制方向"""
    HORIZONTAL = "horizontal"  # 横向（向右）
    VERTICAL = "vertical"     # 纵向（向下）


class WriteMode(str, Enum):
    """写入模式"""
    SKIP_NONEMPTY = "skip_nonempty"  # 跳过非空单元格
    OVERWRITE = "overwrite"          # 直接覆盖（预留接口）


class MatchMode(str, Enum):
    """Sheet 匹配模式"""
    EXACT = "exact"      # 精确匹配
    FUZZY = "fuzzy"      # 模糊匹配（包含关系）
    MAPPING = "mapping"  # 用户自定义映射
