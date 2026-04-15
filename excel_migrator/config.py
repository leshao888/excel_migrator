"""全局配置"""
import os
import sys
from pathlib import Path

from core.enums import WriteMode

# 项目根目录（开发时指向源码，打包后指向临时解压目录）
PROJECT_ROOT = Path(__file__).parent

# 判断是否为打包环境
IS_FROZEN = getattr(sys, 'frozen', False)

# 存储目录
# 开发模式：使用项目目录下的 storage
# 打包模式：使用用户数据目录（%APPDATA%/ExcelMigrator/storage），确保数据持久化
if IS_FROZEN:
    # 打包环境：使用用户目录，确保数据持久化
    STORAGE_DIR = Path(os.environ.get('APPDATA', os.path.expanduser('~'))) / 'ExcelMigrator' / 'storage'
else:
    # 开发环境：使用项目目录
    STORAGE_DIR = PROJECT_ROOT / "storage"

STORAGE_DIR.mkdir(parents=True, exist_ok=True)

# 存储文件路径
DATA_SOURCES_FILE = STORAGE_DIR / "data_sources.json"
TEMPLATES_FILE = STORAGE_DIR / "templates.json"
MAPPING_CONFIGS_FILE = STORAGE_DIR / "mapping_configs.json"

# 默认配置
CONFIG_VERSION = "1.3"  # 配置版本，用于后续兼容性扩展
DEFAULT_DIRECTION = "horizontal"
DEFAULT_WRITE_MODE = WriteMode.SKIP_NONEMPTY

# Excel 相关
MAX_SHEET_NAME_LENGTH = 31
SUPPORTED_EXTENSIONS = [".xlsx", ".xlsm"]

# 示例文件
EXAMPLE_MAPPING_FILENAME = "[示例]映射配置.json"
