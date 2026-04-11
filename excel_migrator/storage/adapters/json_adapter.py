"""JSON 存储适配器"""
import json
from pathlib import Path
from typing import Any


class JsonStorageAdapter:
    """JSON 文件存储适配器"""

    def __init__(self, storage_dir: str):
        self._storage_dir = Path(storage_dir)
        self._storage_dir.mkdir(exist_ok=True)

    def _get_file_path(self, filename: str) -> Path:
        """获取文件路径"""
        return self._storage_dir / filename

    def load(self, filename: str, default: Any = None) -> Any:
        """加载 JSON 数据"""
        file_path = self._get_file_path(filename)
        if not file_path.exists():
            return default

        try:
            with open(file_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            return default

    def save(self, filename: str, data: Any) -> None:
        """保存 JSON 数据"""
        file_path = self._get_file_path(filename)
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
