"""存储管理器"""
from typing import Optional

from config import DATA_SOURCES_FILE, TEMPLATES_FILE, MAPPING_CONFIGS_FILE
from core.models import DataSourceItem, TemplateItem, MappingConfigItem, MappingConfig
from storage.adapters.json_adapter import JsonStorageAdapter


class StorageManager:
    """存储管理器"""

    def __init__(self, storage_dir: str):
        self._adapter = JsonStorageAdapter(storage_dir)

    # ==================== DataSource ====================

    def load_data_sources(self) -> list[DataSourceItem]:
        """加载数据源列表"""
        data = self._adapter.load("data_sources.json", {"items": []})
        return [DataSourceItem.from_dict(item) for item in data.get("items", [])]

    def save_data_source(self, item: DataSourceItem) -> bool:
        """保存数据源（基于文件名判断重复），返回是否新增"""
        items = self.load_data_sources()

        # 获取文件名（去除空格）作为唯一标识
        new_filename = item.name.replace(" ", "").replace("\u3000", "")

        # 检查是否已存在（基于文件名去除空格后比较）
        existing_index = None
        for i, existing in enumerate(items):
            existing_filename = existing.name.replace(" ", "").replace("\u3000", "")
            if existing_filename == new_filename:
                existing_index = i
                break

        is_new = True
        if existing_index is not None:
            # 更新时保留原 id
            is_new = False
            item = DataSourceItem(
                id=items[existing_index].id,
                name=item.name,
                file_path=item.file_path,
                sheets=item.sheets,
                added_at=items[existing_index].added_at
            )
            items[existing_index] = item
        else:
            items.append(item)

        self._adapter.save("data_sources.json", {"items": [i.to_dict() for i in items]})
        return is_new

    def delete_data_source(self, id: str) -> bool:
        """删除数据源"""
        items = self.load_data_sources()
        original_length = len(items)
        items = [item for item in items if item.id != id]

        if len(items) == original_length:
            return False

        self._adapter.save("data_sources.json", {"items": [i.to_dict() for i in items]})
        return True

    def get_data_source(self, id: str) -> Optional[DataSourceItem]:
        """获取指定数据源"""
        items = self.load_data_sources()
        for item in items:
            if item.id == id:
                return item
        return None

    def update_data_source(self, item: DataSourceItem) -> bool:
        """更新数据源（基于 id）"""
        items = self.load_data_sources()
        for i, existing in enumerate(items):
            if existing.id == item.id:
                items[i] = item
                self._adapter.save("data_sources.json", {"items": [i.to_dict() for i in items]})
                return True
        return False

    # ==================== Template ====================

    def load_templates(self) -> list[TemplateItem]:
        """加载模板列表"""
        data = self._adapter.load("templates.json", {"items": []})
        return [TemplateItem.from_dict(item) for item in data.get("items", [])]

    def save_template(self, item: TemplateItem) -> bool:
        """保存模板（基于文件名判断重复），返回是否新增"""
        items = self.load_templates()

        # 获取文件名（去除空格）作为唯一标识
        new_filename = item.name.replace(" ", "").replace("\u3000", "")

        # 检查是否已存在（基于文件名去除空格后比较）
        existing_index = None
        for i, existing in enumerate(items):
            existing_filename = existing.name.replace(" ", "").replace("\u3000", "")
            if existing_filename == new_filename:
                existing_index = i
                break

        is_new = True
        if existing_index is not None:
            # 更新时保留原 id
            is_new = False
            item = TemplateItem(
                id=items[existing_index].id,
                name=item.name,
                file_path=item.file_path,
                sheets=item.sheets,
                added_at=items[existing_index].added_at
            )
            items[existing_index] = item
        else:
            items.append(item)

        self._adapter.save("templates.json", {"items": [i.to_dict() for i in items]})
        return is_new

    def delete_template(self, id: str) -> bool:
        """删除模板"""
        items = self.load_templates()
        original_length = len(items)
        items = [item for item in items if item.id != id]

        if len(items) == original_length:
            return False

        self._adapter.save("templates.json", {"items": [i.to_dict() for i in items]})
        return True

    def get_template(self, id: str) -> Optional[TemplateItem]:
        """获取指定模板"""
        items = self.load_templates()
        for item in items:
            if item.id == id:
                return item
        return None

    def update_template(self, item: TemplateItem) -> bool:
        """更新模板（基于 id）"""
        items = self.load_templates()
        for i, existing in enumerate(items):
            if existing.id == item.id:
                items[i] = item
                self._adapter.save("templates.json", {"items": [i.to_dict() for i in items]})
                return True
        return False

    # ==================== MappingConfig ====================

    def load_mapping_configs(self) -> list[MappingConfigItem]:
        """加载映射配置列表"""
        data = self._adapter.load("mapping_configs.json", {"items": []})
        return [MappingConfigItem.from_dict(item) for item in data.get("items", [])]

    def save_mapping_config(self, item: MappingConfigItem) -> None:
        """保存映射配置"""
        items = self.load_mapping_configs()

        existing_index = None
        for i, existing in enumerate(items):
            if existing.id == item.id:
                existing_index = i
                break

        if existing_index is not None:
            items[existing_index] = item
        else:
            items.append(item)

        self._adapter.save("mapping_configs.json", {"items": [i.to_dict() for i in items]})

    def delete_mapping_config(self, id: str) -> bool:
        """删除映射配置"""
        items = self.load_mapping_configs()
        original_length = len(items)
        items = [item for item in items if item.id != id]

        if len(items) == original_length:
            return False

        self._adapter.save("mapping_configs.json", {"items": [i.to_dict() for i in items]})
        return True

    def get_mapping_config(self, id: str) -> Optional[MappingConfigItem]:
        """获取指定映射配置"""
        items = self.load_mapping_configs()
        for item in items:
            if item.id == id:
                return item
        return None

    def update_mapping_config(self, item: MappingConfigItem) -> bool:
        """更新映射配置（基于 id）"""
        items = self.load_mapping_configs()
        for i, existing in enumerate(items):
            if existing.id == item.id:
                items[i] = item
                self._adapter.save("mapping_configs.json", {"items": [i.to_dict() for i in items]})
                return True
        return False
