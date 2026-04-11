"""执行迁移页面"""
import os
import shutil
import streamlit as st

from core.copier import DataCopier
from core.enums import WriteMode
from storage.store import StorageManager
from ui.components import render_result_message, try_load_excel
from utils.excel_utils import get_sheet_names
from utils.path_utils import file_exists, is_file_locked


def _generate_copy_path(original_path: str) -> str:
    """生成副本文件路径（在同一目录）"""
    directory = os.path.dirname(original_path)
    filename = os.path.basename(original_path)
    name, ext = os.path.splitext(filename)
    copy_filename = f"{name}_副本{ext}"
    return os.path.join(directory, copy_filename)


def render_page(store: StorageManager):
    """渲染执行迁移页面"""
    st.header("执行迁移")

    # 加载数据源和模板
    data_sources = store.load_data_sources()
    templates = store.load_templates()
    configs = store.load_mapping_configs()

    # 选择数据源
    st.subheader("选择数据源")
    if not data_sources:
        st.error("暂无数据源，请先在【数据源管理】中添加")
        return

    ds_options = {ds.id: f"{ds.name}" for ds in data_sources}
    selected_ds_id = st.selectbox(
        "数据源",
        options=list(ds_options.keys()),
        format_func=lambda x: ds_options[x],
        key="migrate_ds_selector"
    )
    selected_ds = next((ds for ds in data_sources if ds.id == selected_ds_id), None)

    # 验证数据源文件
    if selected_ds:
        if not file_exists(selected_ds.file_path):
            st.error(f"⚠️ 数据源文件不存在: {selected_ds.file_path}")
            return
        st.info(f"路径: `{selected_ds.file_path}`")

    # 选择模板
    st.subheader("选择模板（将被修改）")
    if not templates:
        st.error("暂无模板，请先在【模板管理】中添加")
        return

    t_options = {t.id: f"{t.name}" for t in templates}
    selected_t_id = st.selectbox(
        "模板",
        options=list(t_options.keys()),
        format_func=lambda x: t_options[x],
        key="migrate_t_selector"
    )
    selected_t = next((t for t in templates if t.id == selected_t_id), None)

    # 验证模板文件
    if selected_t:
        if not file_exists(selected_t.file_path):
            st.error(f"⚠️ 模板文件不存在: {selected_t.file_path}")
            return
        st.info(f"路径: `{selected_t.file_path}`")

    # 选择映射配置
    st.subheader("选择映射配置")
    if not configs:
        st.error("暂无映射配置，请先在【映射配置】中添加")
        return

    config_options = {c.id: c.name for c in configs}
    selected_config_id = st.selectbox(
        "映射配置",
        options=list(config_options.keys()),
        format_func=lambda x: config_options[x],
        key="migrate_config_selector"
    )
    selected_config_item = next((c for c in configs if c.id == selected_config_id), None)

    if not selected_config_item:
        return

    config = selected_config_item.config

    # 构建模板名称到 TemplateItem 的映射
    template_name_map = {os.path.basename(t.file_path): t for t in templates}

    # 显示将要执行的映射预览
    st.divider()
    st.subheader("映射预览")

    matched_sheets = []
    unmatched_source = []
    unmatched_target = []

    for mapping in config.sheet_mappings:
        ds_sheet = mapping.data_source_sheet
        t_file = mapping.template_file

        # 检查数据源 sheet 是否存在
        if ds_sheet not in selected_ds.sheets:
            unmatched_source.append(ds_sheet)
            continue

        # 检查模板文件是否在已加载的模板中
        if t_file not in template_name_map:
            unmatched_target.append(t_file)
            continue

        matched_sheets.append(mapping)

    if matched_sheets:
        st.success(f"✅ 将迁移 {len(matched_sheets)} 个 Sheet")
        for m in matched_sheets:
            rule = m.copy_rule
            if rule.is_array_mode():
                # 数组模式：显示每个配对
                pairs = rule.get_pairs()
                st.markdown(f"- **{m.data_source_sheet}** → {m.template_file} ({rule.direction.value}, {rule.length}格)")
                for j, (src, tgt) in enumerate(pairs):
                    st.markdown(f"  {j+1}. {src} → {tgt}")
            else:
                st.markdown(f"- {m.data_source_sheet} → {m.template_file} ({rule.source_start} → {rule.target_start}, {rule.direction.value}, {rule.length}格)")

    if unmatched_source:
        st.warning(f"⚠️ 数据源中不存在 {len(unmatched_source)} 个 Sheet")
        for sheet in unmatched_source:
            st.markdown(f"- {sheet}")

    if unmatched_target:
        st.warning(f"⚠️ 未找到对应的模板文件 {len(unmatched_target)} 个")
        for fname in unmatched_target:
            st.markdown(f"- {fname}")

    # 执行迁移
    st.divider()

    if st.button("🚀 执行迁移", key="execute_migrate", type="primary"):
        if not selected_ds or not selected_t or not selected_config_item:
            st.error("请确保已选择数据源、模板和映射配置")
            return

        try:
            # 检查数据源文件
            if is_file_locked(selected_ds.file_path):
                st.error("⚠️ 数据源文件已打开，请关闭后重试")
                return
            if is_file_locked(selected_t.file_path):
                st.error("⚠️ 模板文件已打开，请关闭后重试")
                return

            # 加载数据源文件（使用 data_only=True 读取公式的计算值）
            source_wb, error = try_load_excel(selected_ds.file_path, data_only=True)
            if error:
                st.error(f"❌ 加载数据源失败: {error}")
                return

            # 按模板文件分组执行迁移
            all_results = []
            copied_files = []  # 记录生成副本的文件路径

            # 按 template_file 分组
            template_groups = {}
            for mapping in config.sheet_mappings:
                t_file = mapping.template_file
                if t_file not in template_groups:
                    template_groups[t_file] = []
                template_groups[t_file].append(mapping)

            for t_file, mappings in template_groups.items():
                # 查找对应的模板
                template_item = template_name_map.get(t_file)
                if not template_item:
                    continue

                original_path = template_item.file_path
                target_path = original_path

                # 检查模板文件是否被占用
                if is_file_locked(original_path):
                    target_path = _generate_copy_path(original_path)
                    shutil.copy2(original_path, target_path)
                    copied_files.append(target_path)

                # 加载模板文件
                target_wb, load_error = try_load_excel(target_path)
                if load_error:
                    st.error(f"❌ 加载模板失败: {load_error}")
                    return

                # 执行复制（使用覆盖模式，确保数据正确写入）
                copier = DataCopier(write_mode=WriteMode.OVERWRITE)
                for mapping in mappings:
                    # 只复制数据源 sheet 在选择的数据源中存在的映射
                    if mapping.data_source_sheet not in selected_ds.sheets:
                        continue
                    result = copier.copy(source_wb, target_wb, mapping)
                    all_results.append(result)

                # 保存模板文件
                target_wb.save(target_path)

            # 显示结果
            st.divider()
            st.subheader("迁移结果")
            render_result_message(all_results)

            # 显示副本文件信息
            if copied_files:
                for copy_path in copied_files:
                    st.info(f"📁 原始文件被占用，已生成副本: `{copy_path}`")

        except PermissionError:
            st.error("⚠️ 文件已打开，请关闭后重试")
        except Exception as e:
            st.error(f"迁移失败: {str(e)}")
