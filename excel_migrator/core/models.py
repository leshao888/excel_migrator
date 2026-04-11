"""数据模型定义"""
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional, Union
import uuid

from core.enums import Direction, WriteMode


@dataclass
class CopyRule:
    """复制规则"""
    source_start: Union[str, list[str]]  # 数据源起始单元格，如 "A1" 或 ["A1", "B1", "C1"]
    target_start: Union[str, list[str]]   # 模板起始单元格，如 "B3" 或 ["D1", "E1", "F1"]
    direction: Direction                  # 方向
    length: int                          # 长度

    @classmethod
    def from_dict(cls, data: dict) -> "CopyRule":
        return cls(
            source_start=data["source_start"],
            target_start=data["target_start"],
            direction=Direction(data.get("direction", "horizontal")),
            length=data["length"]
        )

    def to_dict(self) -> dict:
        return {
            "source_start": self.source_start,
            "target_start": self.target_start,
            "direction": self.direction.value,
            "length": self.length
        }

    def is_array_mode(self) -> bool:
        """判断是否为数组模式"""
        return isinstance(self.source_start, list) or isinstance(self.target_start, list)

    def get_pairs(self) -> list:
        """获取源和目标的配对列表"""
        if isinstance(self.source_start, str):
            source_list = [self.source_start]
        else:
            source_list = self.source_start

        if isinstance(self.target_start, str):
            target_list = [self.target_start]
        else:
            target_list = self.target_start

        # 取最短长度
        min_len = min(len(source_list), len(target_list))
        return list(zip(source_list[:min_len], target_list[:min_len]))


@dataclass
class SheetMapping:
    """Sheet 映射配置"""
    data_source_sheet: str   # 数据源的 sheet 名称
    template_file: str        # 目标模板文件名称
    template_sheet: str = ""  # 目标模板的 sheet 名称（可选，默认使用第一个）
    copy_rule: CopyRule = None  # 修复：添加默认值

    def __post_init__(self):
        if self.copy_rule is None:
            self.copy_rule = CopyRule(
                source_start="A1",
                target_start="A1",
                direction=Direction.HORIZONTAL,
                length=10
            )

    @classmethod
    def from_dict(cls, data: dict) -> "SheetMapping":
        # 兼容旧字段名
        source_sheet = data.get("data_source_sheet") or data.get("source_sheet", "")
        target = data.get("template_file") or data.get("target_sheet", "")
        tmpl_sheet = data.get("template_sheet", "")  # 兼容旧版本
        return cls(
            data_source_sheet=source_sheet,
            template_file=target,
            template_sheet=tmpl_sheet,
            copy_rule=CopyRule.from_dict(data["copy_rule"])
        )

    def to_dict(self) -> dict:
        result = {
            "data_source_sheet": self.data_source_sheet,
            "template_file": self.template_file,
            "copy_rule": self.copy_rule.to_dict()
        }
        if self.template_sheet:
            result["template_sheet"] = self.template_sheet
        return result


@dataclass
class MappingConfig:
    """完整映射配置"""
    version: str = "1.0"
    default_direction: Direction = Direction.HORIZONTAL
    default_write_mode: WriteMode = WriteMode.SKIP_NONEMPTY
    sheet_mappings: list[SheetMapping] = field(default_factory=list)

    @classmethod
    def from_dict(cls, data: dict) -> "MappingConfig":
        sheet_mappings = [SheetMapping.from_dict(m) for m in data.get("sheet_mappings", [])]
        return cls(
            version=data.get("version", "1.0"),
            default_direction=Direction(data.get("default_direction", "horizontal")),
            default_write_mode=WriteMode(data.get("default_write_mode", "skip_nonempty")),
            sheet_mappings=sheet_mappings
        )

    def to_dict(self) -> dict:
        return {
            "version": self.version,
            "default_direction": self.default_direction.value,
            "default_write_mode": self.default_write_mode.value,
            "sheet_mappings": [m.to_dict() for m in self.sheet_mappings]
        }


@dataclass
class DataSourceItem:
    """数据源条目"""
    id: str
    name: str
    file_path: str
    sheets: list[str]
    added_at: str

    @classmethod
    def create(cls, name: str, file_path: str, sheets: list[str]) -> "DataSourceItem":
        return cls(
            id=str(uuid.uuid4()),
            name=name,
            file_path=file_path,
            sheets=sheets,
            added_at=datetime.now().isoformat()
        )

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "name": self.name,
            "file_path": self.file_path,
            "sheets": self.sheets,
            "added_at": self.added_at
        }

    @classmethod
    def from_dict(cls, data: dict) -> "DataSourceItem":
        return cls(**data)


@dataclass
class TemplateItem:
    """模板条目"""
    id: str
    name: str
    file_path: str
    sheets: list[str]
    added_at: str

    @classmethod
    def create(cls, name: str, file_path: str, sheets: list[str]) -> "TemplateItem":
        return cls(
            id=str(uuid.uuid4()),
            name=name,
            file_path=file_path,
            sheets=sheets,
            added_at=datetime.now().isoformat()
        )

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "name": self.name,
            "file_path": self.file_path,
            "sheets": self.sheets,
            "added_at": self.added_at
        }

    @classmethod
    def from_dict(cls, data: dict) -> "TemplateItem":
        return cls(**data)


@dataclass
class MappingConfigItem:
    """映射配置条目"""
    id: str
    name: str
    config: MappingConfig
    created_at: str
    updated_at: str

    @classmethod
    def create(cls, name: str, config: MappingConfig) -> "MappingConfigItem":
        now = datetime.now().isoformat()
        return cls(
            id=str(uuid.uuid4()),
            name=name,
            config=config,
            created_at=now,
            updated_at=now
        )

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "name": self.name,
            "config": self.config.to_dict(),
            "created_at": self.created_at,
            "updated_at": self.updated_at
        }

    @classmethod
    def from_dict(cls, data: dict) -> "MappingConfigItem":
        return cls(
            id=data["id"],
            name=data["name"],
            config=MappingConfig.from_dict(data["config"]),
            created_at=data["created_at"],
            updated_at=data["updated_at"]
        )


@dataclass
class CopyResult:
    """复制结果"""
    success: bool
    source_sheet: str
    template_file: str
    message: str
    rows_copied: int = 0
