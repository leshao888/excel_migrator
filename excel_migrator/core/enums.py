"""枚举定义"""
import re
from enum import Enum
from typing import List, Tuple


# ============ 常量定义 ============

# 分隔符类型
SEPARATOR_CHINESE_COMMA = "，"  # 中文逗号
SEPARATOR_ENGLISH_COMMA = ","   # 英文逗号
SEPARATOR_CHINESE_PAUSE = "、"  # 中文顿号
VALID_SEPARATORS = {SEPARATOR_ENGLISH_COMMA, SEPARATOR_CHINESE_COMMA, SEPARATOR_CHINESE_PAUSE}
SEPARATOR_PATTERN = re.compile(r"[，,、\s]+")


def normalize_string(value: str) -> str:
    """规范化字符串：全角转半角，大小写转小写，中文逗号/顿号转英文逗号"""
    if not isinstance(value, str):
        return value
    result = ""
    for char in value:
        code = ord(char)
        if code == 0x3000:  # 全角空格
            result += " "
        elif 0xFF01 <= code <= 0xFF5E:  # 全角字符
            result += chr(code - 0xFEE0)
        else:
            result += char
    # 中文逗号/顿号转英文逗号
    result = result.replace(SEPARATOR_CHINESE_COMMA, SEPARATOR_ENGLISH_COMMA)
    result = result.replace(SEPARATOR_CHINESE_PAUSE, SEPARATOR_ENGLISH_COMMA)
    # 转小写
    result = result.lower()
    return result.strip()


def normalize_list_string(value: str) -> str:
    """规范化列表字符串（如 "A1, B2, C3" 或 "A1，B2，C3" 或 "A1、B2、C3"）"""
    if not isinstance(value, str):
        return value
    # 先规范化字符串（处理全角转半角、中文逗号/顿号）
    value = normalize_string(value)
    # 使用分隔符分割
    parts = SEPARATOR_PATTERN.split(value)
    return ",".join(part.strip() for part in parts if part.strip())


def validate_cell_references(value: str) -> Tuple[bool, str]:
    """
    验证单元格引用格式是否合法
    Returns: (is_valid, error_message)
    """
    if not value or not value.strip():
        return False, "单元格引用不能为空"

    normalized = normalize_string(value)
    # 检查是否只包含合法的分隔符
    parts = SEPARATOR_PATTERN.split(normalized)

    cell_pattern = re.compile(r'^[A-Za-z]+\d+$')
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if not cell_pattern.match(part):
            return False, f"无效的单元格引用: '{part}'，应为如 A1, B2, AA123 等格式"

    return True, ""


def get_supported_separators_display() -> str:
    """获取支持的分割符显示文本"""
    return "英文逗号 (,)、中文逗号 (，)、中文顿号 (、)"


class Direction(str, Enum):
    """数据复制方向"""
    HORIZONTAL = "horizontal"  # 横向（向右）
    VERTICAL = "vertical"     # 纵向（向下）

    @classmethod
    def from_string(cls, value: str) -> "Direction":
        """从字符串创建枚举，支持大小写不敏感和全角"""
        if not value:
            return cls.HORIZONTAL
        normalized = normalize_string(value)
        try:
            return cls(normalized)
        except ValueError:
            # 尝试大小写不敏感查找
            for member in cls:
                if member.value == normalized:
                    return member
            return cls.HORIZONTAL  # 默认值


class WriteMode(str, Enum):
    """写入模式"""
    SKIP_NONEMPTY = "skip_nonempty"  # 跳过非空单元格（只写空行）
    OVERWRITE = "overwrite"           # 直接覆盖

    @classmethod
    def from_string(cls, value: str) -> "WriteMode":
        """从字符串创建枚举，支持大小写不敏感和全角"""
        if not value:
            return cls.SKIP_NONEMPTY
        normalized = normalize_string(value)
        try:
            return cls(normalized)
        except ValueError:
            # 尝试大小写不敏感查找
            for member in cls:
                if member.value == normalized:
                    return member
            return cls.SKIP_NONEMPTY  # 默认值


class WriteModeHandler:
    """写入模式处理器接口（策略模式）"""

    @staticmethod
    def should_write(cell_value) -> bool:
        """判断是否应该写入"""
        raise NotImplementedError


class SkipNonemptyHandler(WriteModeHandler):
    """跳过非空处理器"""

    @staticmethod
    def should_write(cell_value) -> bool:
        return cell_value is None or str(cell_value).startswith('=')


class OverwriteHandler(WriteModeHandler):
    """覆盖写入处理器"""

    @staticmethod
    def should_write(cell_value) -> bool:
        return True


# 写入模式处理器注册表（便于扩展）
WRITE_MODE_HANDLERS = {
    WriteMode.SKIP_NONEMPTY: SkipNonemptyHandler,
    WriteMode.OVERWRITE: OverwriteHandler,
}


def get_write_mode_handler(mode: WriteMode) -> WriteModeHandler:
    """获取写入模式处理器"""
    handler = WRITE_MODE_HANDLERS.get(mode)
    if handler is None:
        raise ValueError(f"未知的写入模式: {mode}")
    return handler


class MatchMode(str, Enum):
    """Sheet 匹配模式"""
    EXACT = "exact"      # 精确匹配
    FUZZY = "fuzzy"      # 模糊匹配（包含关系）
    MAPPING = "mapping"  # 用户自定义映射
