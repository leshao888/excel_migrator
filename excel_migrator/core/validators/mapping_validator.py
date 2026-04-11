"""映射配置验证器"""
from dataclasses import dataclass
from typing import Optional

from openpyxl import Workbook

from core.models import MappingConfig, SheetMapping
from core.enums import Direction


@dataclass
class MappingValidationError:
    """映射验证错误"""
    sheet: str
    field: str
    message: str


class MappingValidator:
    """映射验证器"""

    def validate(
        self,
        config: MappingConfig,
        source_sheets: list[str],
        template_files: list[str]
    ) -> list[MappingValidationError]:
        """验证映射配置"""
        errors = []

        for mapping in config.sheet_mappings:
            # 验证 data_source_sheet 是否存在
            if mapping.data_source_sheet not in source_sheets:
                errors.append(MappingValidationError(
                    sheet=mapping.data_source_sheet,
                    field="data_source_sheet",
                    message=f"数据源中不存在此 Sheet"
                ))

            # 验证 template_file 是否存在（通过文件名匹配）
            if mapping.template_file not in template_files:
                errors.append(MappingValidationError(
                    sheet=mapping.template_file,
                    field="template_file",
                    message=f"未找到对应的模板文件"
                ))

            # 验证复制规则
            rule_errors = self._validate_copy_rule(mapping)
            errors.extend(rule_errors)

        return errors

    def _validate_copy_rule(self, mapping: SheetMapping) -> list[MappingValidationError]:
        """验证复制规则"""
        errors = []
        rule = mapping.copy_rule

        # 验证起始单元格格式
        if not self._is_valid_cell(rule.source_start):
            errors.append(MappingValidationError(
                sheet=mapping.data_source_sheet,
                field="source_start",
                message=f"无效的单元格地址: {rule.source_start}"
            ))

        if not self._is_valid_cell(rule.target_start):
            errors.append(MappingValidationError(
                sheet=mapping.template_file,
                field="target_start",
                message=f"无效的单元格地址: {rule.target_start}"
            ))

        # 验证长度
        if rule.length <= 0:
            errors.append(MappingValidationError(
                sheet=mapping.data_source_sheet,
                field="length",
                message=f"长度必须大于 0"
            ))

        return errors

    def _is_valid_cell(self, cell: str) -> bool:
        """验证单元格地址格式"""
        import re
        pattern = r'^[A-Z]+\d+$'
        return bool(re.match(pattern, cell.upper()))
