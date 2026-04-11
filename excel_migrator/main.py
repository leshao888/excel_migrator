"""Excel 数据迁移工具 - 主入口"""
import streamlit as st

from config import STORAGE_DIR
from storage.store import StorageManager
from ui.pages.data_source_page import render_page as render_data_source
from ui.pages.template_page import render_page as render_template
from ui.pages.mapping_page import render_page as render_mapping
from ui.pages.migrate_page import render_page as render_migrate


def main():
    """主函数"""
    st.set_page_config(
        page_title="Excel 数据迁移工具",
        page_icon="📊",
        layout="wide"
    )

    # 每次都创建新的 StorageManager，确保从文件重新读取
    store = StorageManager(str(STORAGE_DIR))

    # 侧边栏导航
    st.sidebar.title("导航")
    page = st.sidebar.radio(
        "选择功能",
        ["数据源管理", "模板管理", "映射配置", "执行迁移"],
        index=0
    )

    st.sidebar.divider()
    st.sidebar.markdown("**Excel 数据迁移工具 v1.0**")

    # 渲染页面
    if page == "数据源管理":
        render_data_source(store)
    elif page == "模板管理":
        render_template(store)
    elif page == "映射配置":
        render_mapping(store)
    elif page == "执行迁移":
        render_migrate(store)


if __name__ == "__main__":
    main()
