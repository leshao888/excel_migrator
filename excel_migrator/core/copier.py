"""数据复制核心引擎"""
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from core.enums import Direction, WriteMode
from core.models import CopyResult, MappingConfig, SheetMapping
from utils.excel_utils import (
    read_range_data_with_format,
    write_range_data_with_format,
    get_worksheet,
    get_sheet_names
)


class DataCopier:
    """数据复制器"""

    def __init__(self, write_mode: WriteMode = WriteMode.SKIP_NONEMPTY):
        self._write_mode = write_mode

    def copy(
        self,
        source_wb: Workbook,
        target_wb: Workbook,
        mapping: SheetMapping,
        target_sheet_name: str = None
    ) -> CopyResult:
        """执行单个 sheet 的数据复制（保留格式）"""
        try:
            # 数据源 sheet
            source_sheet_name = mapping.data_source_sheet
            # 目标模板文件名称（用于显示）
            template_file = mapping.template_file
            # 目标 sheet 名称（模板的默认 sheet）
            if target_sheet_name is None:
                target_sheet_name = get_sheet_names(target_wb)[0]  # 使用模板的第一个 sheet

            if source_sheet_name not in get_sheet_names(source_wb):
                return CopyResult(
                    success=False,
                    source_sheet=source_sheet_name,
                    template_file=template_file,
                    message=f"数据源中不存在 Sheet: {source_sheet_name}"
                )

            source_ws = get_worksheet(source_wb, source_sheet_name)
            target_ws = get_worksheet(target_wb, target_sheet_name)

            rule = mapping.copy_rule

            # 处理数组模式
            if rule.is_array_mode():
                pairs = rule.get_pairs()
                total_rows = 0
                for src_start, tgt_start in pairs:
                    # 读取数据（包含格式）
                    data, formats, types = read_range_data_with_format(
                        source_ws,
                        src_start,
                        rule.direction,
                        rule.length
                    )
                    # 写入数据（保留格式）
                    write_range_data_with_format(
                        target_ws,
                        tgt_start,
                        rule.direction,
                        rule.length,
                        data,
                        formats,
                        types,
                        self._write_mode.value
                    )
                    total_rows += len(data)
                return CopyResult(
                    success=True,
                    source_sheet=source_sheet_name,
                    template_file=template_file,
                    message=f"成功复制 {len(pairs)} 组数据到 {template_file} ({target_sheet_name})",
                    rows_copied=total_rows
                )
            else:
                # 单值模式
                data, formats, types = read_range_data_with_format(
                    source_ws,
                    rule.source_start,
                    rule.direction,
                    rule.length
                )

                write_range_data_with_format(
                    target_ws,
                    rule.target_start,
                    rule.direction,
                    rule.length,
                    data,
                    formats,
                    types,
                    self._write_mode.value
                )

                return CopyResult(
                    success=True,
                    source_sheet=source_sheet_name,
                    template_file=template_file,
                    message=f"成功复制到 {template_file} ({target_sheet_name})",
                    rows_copied=len(data)
                )

        except Exception as e:
            return CopyResult(
                success=False,
                source_sheet=mapping.data_source_sheet,
                template_file=mapping.template_file,
                message=f"复制失败: {str(e)}"
            )

    def copy_all(
        self,
        source_wb: Workbook,
        target_wb: Workbook,
        config: MappingConfig
    ) -> list[CopyResult]:
        """执行所有 sheet 的数据复制"""
        results = []
        for mapping in config.sheet_mappings:
            result = self.copy(source_wb, target_wb, mapping)
            results.append(result)
        return results
