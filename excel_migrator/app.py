"""
Excel 数据迁移工具 - 独立运行入口
"""
import sys
import os

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 导入并运行 Streamlit 应用
from main import main

if __name__ == "__main__":
    # 设置 Streamlit 配置
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
    os.environ["STREAMLIT_SERVER_PORT"] = "8501"
    os.environ["STREAMLIT_SERVER_ADDRESS"] = "localhost"

    # 运行应用
    main()
