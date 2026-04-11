"""路径工具函数"""
import os
from pathlib import Path


def normalize_path(path: str) -> str:
    """标准化路径"""
    return os.path.normpath(path)


def is_valid_excel_path(path: str) -> bool:
    """检查是否是有效的 Excel 文件路径"""
    if not path:
        return False

    ext = Path(path).suffix.lower()
    return ext in [".xlsx", ".xlsm"]


def file_exists(path: str) -> bool:
    """检查文件是否存在"""
    return os.path.exists(path)


def get_filename(path: str) -> str:
    """获取文件名（不含路径）"""
    return os.path.basename(path)


def get_file_extension(path: str) -> str:
    """获取文件扩展名"""
    return Path(path).suffix.lower()


def is_file_locked(file_path: str) -> bool:
    """检查文件是否被占用"""
    if not os.path.exists(file_path):
        return False

    try:
        with open(file_path, 'a'):
            pass
        return False
    except IOError:
        return True
