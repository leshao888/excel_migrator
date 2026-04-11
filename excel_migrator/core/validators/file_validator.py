"""文件验证器"""
from dataclasses import dataclass
from typing import Optional

from utils.path_utils import file_exists, is_valid_excel_path


@dataclass
class FileValidationError:
    """文件验证错误"""
    field: str
    message: str


class FileValidator:
    """文件验证器"""

    @staticmethod
    def validate_file_path(file_path: str) -> Optional[FileValidationError]:
        """验证文件路径"""
        if not file_path:
            return FileValidationError(
                field="file_path",
                message="文件路径不能为空"
            )

        if not is_valid_excel_path(file_path):
            return FileValidationError(
                field="file_path",
                message="仅支持 .xlsx 或 .xlsm 格式的 Excel 文件"
            )

        if not file_exists(file_path):
            return FileValidationError(
                field="file_path",
                message="文件不存在，请重新选择"
            )

        return None
