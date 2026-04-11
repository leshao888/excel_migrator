"""数据源管理页面"""
import streamlit as st

from core.models import DataSourceItem
from storage.store import StorageManager
from ui.components import render_file_upload_simple, try_load_excel
from utils.excel_utils import get_sheet_names, load_excel
from utils.path_utils import file_exists, is_file_locked


def render_page(store: StorageManager):
    """渲染数据源管理页面"""
    st.header("数据源管理")

    # 显示配置文件位置
    from config import DATA_SOURCES_FILE
    st.success(f"📁 配置文件位置: `{DATA_SOURCES_FILE}`")

    # 上传文件
    file_path, file_name, error = render_file_upload_simple("数据源", "ds")

    if error:
        if error != "请选择文件":
            st.error(f"❌ {error}")
        # 显示已有数据源列表
        _render_datasource_list(store)
        return

    # 名称输入
    name = st.text_input(
        "数据源名称",
        value=file_name if file_name else "",
        placeholder="为这个数据源起个名字",
        key="ds_name_input"
    )

    col1, col2 = st.columns(2)
    with col1:
        save_btn = st.button("💾 保存", type="primary")
    with col2:
        cancel_btn = st.button("取消")

    if cancel_btn:
        st.rerun()

    if save_btn:
        if not name:
            st.error("❌ 请输入名称")
            return

        if not file_path:
            st.error("❌ 请选择文件")
            return

        # 检查文件
        wb, load_error = try_load_excel(file_path)
        if load_error:
            st.error(f"❌ {load_error}")
            return

        try:
            sheets = get_sheet_names(wb)
            del wb

            # 创建并保存
            item = DataSourceItem.create(name=name, file_path=file_path, sheets=sheets)
            is_new = store.save_data_source(item)

            if is_new:
                st.success(f"✅ 新增数据源 '{name}'，包含 {len(sheets)} 个 Sheet")
            else:
                st.success(f"✅ 更新数据源 '{name}'，包含 {len(sheets)} 个 Sheet")
            st.rerun()

        except Exception as e:
            st.error(f"❌ 处理失败: {str(e)}")

    # 显示已有数据源列表
    st.divider()
    _render_datasource_list(store)


def _render_datasource_list(store: StorageManager):
    """渲染数据源列表"""
    data_sources = store.load_data_sources()
    st.subheader(f"已保存的数据源 ({len(data_sources)} 个)")

    if not data_sources:
        st.info("暂无已保存的数据源")
        return

    for ds in data_sources:
        with st.expander(f"📄 {ds.name}", expanded=False):
            st.markdown(f"**路径:** `{ds.file_path}`")
            st.markdown(f"**Sheet数量:** {len(ds.sheets)}")

            # 检查文件状态
            if not file_exists(ds.file_path):
                st.error("⚠️ 文件不存在，请重新上传")
            elif is_file_locked(ds.file_path):
                st.warning("⚠️ 文件已被其他程序打开")
            else:
                # 检测是否修改
                try:
                    wb = load_excel(ds.file_path)
                    actual_sheets = get_sheet_names(wb)
                    del wb
                    if actual_sheets != ds.sheets:
                        st.warning(f"⚠️ 文件已修改，当前有 {len(actual_sheets)} 个 Sheet")
                    else:
                        st.success("✅ 文件正常")
                except Exception as e:
                    st.error(f"读取文件失败: {str(e)}")

            # 操作按钮
            col1, col2 = st.columns(2)
            with col1:
                if st.button("🔄 更新", key=f"update_ds_{ds.id}"):
                    _update_datasource(store, ds)
            with col2:
                if st.button("🗑️ 删除", key=f"delete_ds_{ds.id}"):
                    store.delete_data_source(ds.id)
                    st.success("已删除")
                    st.rerun()


def _update_datasource(store: StorageManager, ds: DataSourceItem):
    """更新数据源"""
    if not file_exists(ds.file_path):
        st.error("文件不存在，无法更新")
        return

    if is_file_locked(ds.file_path):
        st.error("文件已打开，请关闭后重试")
        return

    try:
        wb, error = try_load_excel(ds.file_path)
        if error:
            st.error(f"更新失败: {error}")
            return

        actual_sheets = get_sheet_names(wb)
        del wb

        updated_item = DataSourceItem(
            id=ds.id,
            name=ds.name,
            file_path=ds.file_path,
            sheets=actual_sheets,
            added_at=ds.added_at
        )
        store.update_data_source(updated_item)
        st.success(f"已更新，当前共 {len(actual_sheets)} 个 Sheet")
        st.rerun()
    except Exception as e:
        st.error(f"更新失败: {str(e)}")
