"""Sheet 匹配器"""
from abc import ABC, abstractmethod

from core.enums import MatchMode


class SheetMatcher(ABC):
    """Sheet 匹配器抽象基类"""

    @abstractmethod
    def match(self, source_sheet: str, target_sheet: str) -> bool:
        """判断 source_sheet 是否能匹配 target_sheet"""
        pass


class ExactMatcher(SheetMatcher):
    """精确匹配"""

    def match(self, source_sheet: str, target_sheet: str) -> bool:
        return source_sheet == target_sheet


class FuzzyMatcher(SheetMatcher):
    """模糊匹配（包含关系）"""

    def match(self, source_sheet: str, target_sheet: str) -> bool:
        return source_sheet in target_sheet or target_sheet in source_sheet


class MappingMatcher(SheetMatcher):
    """用户自定义映射匹配"""

    def __init__(self, mapping_dict: dict[str, str]):
        self._mapping = mapping_dict

    def match(self, source_sheet: str, target_sheet: str) -> bool:
        return self._mapping.get(source_sheet) == target_sheet


class SheetMatcherFactory:
    """Sheet 匹配器工厂"""

    _matchers: dict[MatchMode, type[SheetMatcher]] = {
        MatchMode.EXACT: ExactMatcher,
        MatchMode.FUZZY: FuzzyMatcher,
        MatchMode.MAPPING: MappingMatcher,
    }

    @classmethod
    def create(cls, mode: MatchMode, **kwargs) -> SheetMatcher:
        matcher_class = cls._matchers.get(mode)
        if not matcher_class:
            raise ValueError(f"Unknown match mode: {mode}")
        return matcher_class(**kwargs)
