"""core 模块初始化"""
from core.enums import Direction, WriteMode, MatchMode
from core.models import (
    CopyRule,
    SheetMapping,
    MappingConfig,
    DataSourceItem,
    TemplateItem,
    MappingConfigItem,
    CopyResult,
)

# 注意：DataCopier 不在这里导入，避免循环导入问题
# 如需使用，请直接 from core.copier import DataCopier

__all__ = [
    "Direction",
    "WriteMode",
    "MatchMode",
    "CopyRule",
    "SheetMapping",
    "MappingConfig",
    "DataSourceItem",
    "TemplateItem",
    "MappingConfigItem",
    "CopyResult",
]
