# Excel 数据迁移工具

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

用于将 Excel 文件中的数据批量迁移到模板文件的桌面工具。

## 功能特性

### 核心功能
- **数据源管理** - 上传和管理数据源 Excel 文件
- **模板管理** - 上传和管理目标模板 Excel 文件
- **映射配置** - 通过 JSON 文件定义数据复制规则
- **执行迁移** - 一键批量执行数据迁移

### 亮点功能
- 迁移前预览确认，避免填错
- 支持横向/纵向复制模式
- 数组模式：一对一、多对一映射
- 写入模式可选：跳过已有内容 / 直接覆盖
- 文件占用自动检测，生成副本
- 保留数字格式，自动调整列宽
- 支持中文逗号、中文顿号分隔符

## 界面预览

桌面 UI，基于 tkinter + ttkbootstrap 构建：

- 四个功能模块：数据源管理 / 模板管理 / 映射配置 / 执行迁移
- 映射编辑器支持动态添加/删除映射规则
- 迁移前显示完整预览，确认后执行

## 快速开始

### 环境要求

- Python 3.8+
- Windows 系统（桌面 UI）

### 安装依赖

```bash
pip install -r requirements.txt
```

### 运行方式

**桌面 UI（推荐）：**
```bash
python gui.py
```

**命令行界面：**
```bash
python main.py
```

## 配置文件格式

```json
{
  "version": "1.3",
  "default_direction": "horizontal",
  "write_mode": "skip_nonempty",
  "sheet_mappings": [
    {
      "data_source": "数据源.xlsx",
      "data_source_sheet": "Sheet1",
      "template_file": "模板.xlsx",
      "template_sheet": "目标Sheet",
      "copy_rule": {
        "source_start": ["A1", "B1", "C1"],
        "target_start": ["D1", "E1", "F1"],
        "direction": "horizontal",
        "length": 12
      }
    }
  ]
}
```

### 字段说明

| 字段 | 说明 |
|------|------|
| `data_source` | 数据源文件名（需与导入的数据源匹配）|
| `data_source_sheet` | 数据源中的 Sheet 名称 |
| `template_file` | 模板文件名（需与导入的模板匹配）|
| `template_sheet` | 模板中的目标 Sheet（可选）|
| `source_start` | 数据源起始单元格，支持数组 |
| `target_start` | 模板目标起始单元格，支持数组 |
| `direction` | 复制方向：`horizontal` / `vertical` |
| `length` | 复制单元格数量 |
| `write_mode` | 写入模式：`skip_nonempty` / `overwrite` |

### 数组模式示例

多组数据一一对应复制：

```json
"copy_rule": {
  "source_start": ["F4", "F11", "F12", "F13"],
  "target_start": ["C5", "C15", "C16", "C17"],
  "direction": "horizontal",
  "length": 12
}
```

对应关系：`F4→C5`, `F11→C15`, `F12→C16`, `F13→C17`

### 垂直复制示例

纵向复制多行数据：

```json
"copy_rule": {
  "source_start": "F10",
  "target_start": "C5",
  "direction": "vertical",
  "length": 5
}
```

从 F10 单元格开始，向下复制 5 个单元格到 C5 起始位置。

## 项目结构

```
excel_migrator/
├── core/                    # 核心业务逻辑
│   ├── models.py            # 数据模型定义
│   ├── copier.py            # 数据复制逻辑
│   └── enums.py             # 枚举类型
├── storage/                 # 存储层
│   └── store.py             # 存储管理器
├── ui/                      # 用户界面
│   └── tkinter_app.py       # tkinter 主应用
├── utils/                   # 工具函数
│   ├── excel_utils.py       # Excel 操作工具
│   └── path_utils.py        # 路径工具
├── config.py                # 全局配置
├── gui.py                   # 桌面 UI 入口
├── main.py                  # 命令行入口
├── requirements.txt         # Python 依赖
├── MANUAL.md                # 详细操作手册
├── RELEASE.md               # 版本更新日志
├── VERSION.md               # 版本说明
└── todo.md                  # 待实现功能清单
```

## 技术栈

- **桌面 UI**: tkinter + ttkbootstrap
- **Excel 处理**: openpyxl
- **持久化**: JSON 文件存储
- **Python 版本**: 3.8+

## 文档

- [操作手册](MANUAL.md) - 详细使用说明
- [版本说明](VERSION.md) - 各版本功能详情
- [更新日志](RELEASE.md) - 版本更新记录

## License

MIT License
