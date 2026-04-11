"""全局配置"""
from pathlib import Path

from core.enums import WriteMode

# 项目根目录
PROJECT_ROOT = Path(__file__).parent

# 存储目录
STORAGE_DIR = PROJECT_ROOT / "storage"
STORAGE_DIR.mkdir(exist_ok=True)

# 存储文件路径
DATA_SOURCES_FILE = STORAGE_DIR / "data_sources.json"
TEMPLATES_FILE = STORAGE_DIR / "templates.json"
MAPPING_CONFIGS_FILE = STORAGE_DIR / "mapping_configs.json"

# 默认配置
DEFAULT_CONFIG_VERSION = "1.0"
DEFAULT_DIRECTION = "horizontal"
DEFAULT_WRITE_MODE = WriteMode.SKIP_NONEMPTY

# Excel 相关
MAX_SHEET_NAME_LENGTH = 31
SUPPORTED_EXTENSIONS = [".xlsx", ".xlsm"]

# 示例文件
EXAMPLE_MAPPING_FILENAME = "[示例]映射配置.json"
