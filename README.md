# Excel 数据迁移工具

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

## 快速开始

### 运行方式

双击 `ExcelMigrator.exe` 启动程序

### 配置文件格式

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

## 版本信息

当前版本：v1.3 (2026-04-14)

详细说明请参考项目文档：
https://github.com/leshao888/excel_migrator
