"""Excel 工具函数"""
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet

from core.enums import Direction


def load_excel(file_path: str, data_only: bool = False) -> Workbook:
    """
    加载 Excel 文件
    data_only=True 时读取公式的计算值（而不是公式本身）
    """
    return load_workbook(file_path, data_only=data_only)


def get_sheet_names(wb: Workbook) -> list[str]:
    """获取所有 sheet 名称"""
    return wb.sheetnames


def get_worksheet(wb: Workbook, sheet_name: str) -> Worksheet:
    """获取指定 sheet"""
    return wb[sheet_name]


def calculate_cell_position(start_cell: str, direction: Direction, length: int) -> list[str]:
    """
    根据起始单元格、方向、长度计算所有单元格地址

    例如: A1, horizontal, 3 -> [A1, A2, A3]
          B2, vertical, 4 -> [B2, C2, D2, E2]
    """
    from openpyxl.utils import get_column_letter, column_index_from_string

    col, row = _parse_cell(start_cell)
    cells = []

    for i in range(length):
        if direction == Direction.HORIZONTAL:
            current_col = col + i
            current_row = row
        else:  # VERTICAL
            current_col = col
            current_row = row + i

        cells.append(f"{get_column_letter(current_col)}{current_row}")

    return cells


def _parse_cell(cell: str) -> tuple[int, int]:
    """解析单元格地址，返回 (列号, 行号)"""
    from openpyxl.utils import column_index_from_string

    col_str = ""
    row_str = ""

    for char in cell:
        if char.isalpha():
            col_str += char
        else:
            row_str += char

    return column_index_from_string(col_str), int(row_str)


def read_range_data(ws: Worksheet, start_cell: str, direction: Direction, length: int) -> list[list]:
    """读取指定范围的数据"""
    cells = calculate_cell_position(start_cell, direction, length)

    if direction == Direction.HORIZONTAL:
        # 横向：按行读取
        row_data = {}
        for cell_addr in cells:
            col, row = _parse_cell(cell_addr)
            if row not in row_data:
                row_data[row] = []
            row_data[row].append((col, ws[cell_addr].value))

        # 按行号排序
        result = []
        for row in sorted(row_data.keys()):
            # 按列号排序
            sorted_cells = sorted(row_data[row], key=lambda x: x[0])
            result.append([val for _, val in sorted_cells])
        return result
    else:
        # 纵向：按列读取
        col_data = {}
        for cell_addr in cells:
            col, row = _parse_cell(cell_addr)
            if col not in col_data:
                col_data[col] = []
            col_data[col].append((row, ws[cell_addr].value))

        # 按列号排序
        result = []
        for col in sorted(col_data.keys()):
            # 按行号排序
            sorted_cells = sorted(col_data[col], key=lambda x: x[0])
            result.append([val for _, val in sorted_cells])
        return result


def write_range_data(
    ws: Worksheet,
    start_cell: str,
    direction: Direction,
    length: int,
    data: list[list],
    write_mode: str = "skip_nonempty"
):
    """写入数据到指定范围"""
    cells = calculate_cell_position(start_cell, direction, length)

    if direction == Direction.HORIZONTAL:
        # 横向：按行写入
        row_data = {}
        for i, cell_addr in enumerate(cells):
            col, row = _parse_cell(cell_addr)
            if row not in row_data:
                row_data[row] = {}
            row_data[row][col] = data[0][i] if i < len(data[0]) else None

        for row, cols in row_data.items():
            for col, value in cols.items():
                cell_addr = f"{_get_col_letter(col)}{row}"
                if write_mode == "skip_nonempty" and ws[cell_addr].value is not None:
                    continue
                ws[cell_addr] = value
    else:
        # 纵向：按列写入
        col_data = {}
        for i, cell_addr in enumerate(cells):
            col, row = _parse_cell(cell_addr)
            if col not in col_data:
                col_data[col] = {}
            col_data[col][row] = data[i][0] if i < len(data) else None

        for col, rows in col_data.items():
            for row, value in rows.items():
                cell_addr = f"{_get_col_letter(col)}{row}"
                if write_mode == "skip_nonempty" and ws[cell_addr].value is not None:
                    continue
                ws[cell_addr] = value


def _get_col_letter(col: int) -> str:
    """列号转字母"""
    from openpyxl.utils import get_column_letter
    return get_column_letter(col)


def read_range_data_with_format(ws: Worksheet, start_cell: str, direction: Direction, length: int) -> tuple:
    """
    读取指定范围的数据和格式信息
    返回: (data, formats, types)
    - data: 二维数据列表
    - formats: 每个单元格的 number_format
    - types: 每个单元格的数据类型
    """
    cells = calculate_cell_position(start_cell, direction, length)

    if direction == Direction.HORIZONTAL:
        # 横向：按行读取
        row_data = {}
        row_formats = {}
        row_types = {}
        for cell_addr in cells:
            col, row = _parse_cell(cell_addr)
            if row not in row_data:
                row_data[row] = []
                row_formats[row] = []
                row_types[row] = []
            cell = ws[cell_addr]
            row_data[row].append((col, cell.value))
            row_formats[row].append((col, cell.number_format))
            row_types[row].append((col, cell.data_type))

        # 按行号排序并整理数据
        result = []
        formats_result = []
        types_result = []
        for row in sorted(row_data.keys()):
            sorted_data = sorted(row_data[row], key=lambda x: x[0])
            sorted_formats = sorted(row_formats[row], key=lambda x: x[0])
            sorted_types = sorted(row_types[row], key=lambda x: x[0])
            result.append([val for _, val in sorted_data])
            formats_result.append([fmt for _, fmt in sorted_formats])
            types_result.append([t for _, t in sorted_types])
        return result, formats_result, types_result
    else:
        # 纵向：按列读取
        col_data = {}
        col_formats = {}
        col_types = {}
        for cell_addr in cells:
            col, row = _parse_cell(cell_addr)
            if col not in col_data:
                col_data[col] = []
                col_formats[col] = []
                col_types[col] = []
            cell = ws[cell_addr]
            col_data[col].append((row, cell.value))
            col_formats[col].append((row, cell.number_format))
            col_types[col].append((row, cell.data_type))

        # 按列号排序并整理数据
        result = []
        formats_result = []
        types_result = []
        for col in sorted(col_data.keys()):
            sorted_data = sorted(col_data[col], key=lambda x: x[0])
            sorted_formats = sorted(col_formats[col], key=lambda x: x[0])
            sorted_types = sorted(col_types[col], key=lambda x: x[0])
            result.append([val for _, val in sorted_data])
            formats_result.append([fmt for _, fmt in sorted_formats])
            types_result.append([t for _, t in sorted_types])
        return result, formats_result, types_result


def write_range_data_with_format(
    ws: Worksheet,
    start_cell: str,
    direction: Direction,
    length: int,
    data: list[list],
    formats: list[list] = None,
    types: list[list] = None,
    write_mode: str = "skip_nonempty"
):
    """
    写入数据到指定范围，同时写入格式
    """
    cells = calculate_cell_position(start_cell, direction, length)

    if direction == Direction.HORIZONTAL:
        # 横向：按行写入
        row_data = {}
        row_formats = {}
        row_types = {}
        for i, cell_addr in enumerate(cells):
            col, row = _parse_cell(cell_addr)
            if row not in row_data:
                row_data[row] = {}
                row_formats[row] = {}
                row_types[row] = {}
            data_row = i // length
            data_col = i % length
            row_data[row][col] = data[data_row][data_col] if data_row < len(data) and data_col < len(data[data_row]) else None
            if formats:
                row_formats[row][col] = formats[data_row][data_col] if data_row < len(formats) and data_col < len(formats[data_row]) else 'General'
            if types:
                row_types[row][col] = types[data_row][data_col] if data_row < len(types) and data_col < len(types[data_row]) else 'n'

        for row, cols in row_data.items():
            for col, value in cols.items():
                cell_addr = f"{_get_col_letter(col)}{row}"
                # 跳过非空单元格（仅检查值，不检查公式）
                cell_value = ws[cell_addr].value
                if write_mode == "skip_nonempty" and cell_value is not None and not str(cell_value).startswith('='):
                    continue
                ws[cell_addr] = value
                # 应用格式
                if formats and row in row_formats and col in row_formats[row]:
                    ws[cell_addr].number_format = row_formats[row][col]
    else:
        # 纵向：按列写入
        col_data = {}
        col_formats = {}
        col_types = {}
        for i, cell_addr in enumerate(cells):
            col, row = _parse_cell(cell_addr)
            if col not in col_data:
                col_data[col] = {}
                col_formats[col] = {}
                col_types[col] = {}
            data_row = i % length
            data_col = i // length
            col_data[col][row] = data[data_row][data_col] if data_row < len(data) and data_col < len(data[data_row]) else None
            if formats:
                col_formats[col][row] = formats[data_row][data_col] if data_row < len(formats) and data_col < len(formats[data_row]) else 'General'
            if types:
                col_types[col][row] = types[data_row][data_col] if data_row < len(types) and data_col < len(types[data_row]) else 'n'

        for col, rows in col_data.items():
            for row, value in rows.items():
                cell_addr = f"{_get_col_letter(col)}{row}"
                # 跳过非空单元格（仅检查值，不检查公式）
                cell_value = ws[cell_addr].value
                if write_mode == "skip_nonempty" and cell_value is not None and not str(cell_value).startswith('='):
                    continue
                ws[cell_addr] = value
                # 应用格式
                if formats and col in col_formats and row in col_formats[col]:
                    ws[cell_addr].number_format = col_formats[col][row]
