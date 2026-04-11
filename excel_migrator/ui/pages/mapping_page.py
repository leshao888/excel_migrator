"""映射配置页面"""
import json
import os
from datetime import datetime
import streamlit as st

from config import EXAMPLE_MAPPING_FILENAME, DEFAULT_DIRECTION, DEFAULT_WRITE_MODE
from core.models import MappingConfig, MappingConfigItem, SheetMapping, CopyRule
from core.enums import Direction, WriteMode
from storage.store import StorageManager


def _create_example_config() -> dict:
    """创建示例配置"""
    return {
        "version": "1.0",
        "default_direction": DEFAULT_DIRECTION,
        "default_write_mode": DEFAULT_WRITE_MODE.value,
        "sheet_mappings": [
            {
                "data_source_sheet": "数据源Sheet名称",
                "template_file": "模板文件名.xlsx",
                "copy_rule": {
                    "source_start": ["A1", "B1", "C1"],
                    "target_start": ["D1", "E1", "F1"],
                    "direction": "horizontal",
                    "length": 12
                }
            }
        ]
    }


def render_page(store: StorageManager):
    """渲染映射配置页面"""
    st.header("映射配置")

    # 下载示例文件
    col1, col2 = st.columns([1, 3])
    with col1:
        example_config = _create_example_config()
        st.download_button(
            "📥 下载示例文件",
            data=json.dumps(example_config, ensure_ascii=False, indent=2),
            file_name=EXAMPLE_MAPPING_FILENAME,
            mime="application/json",
            key="download_example"
        )

    st.info(f"请下载示例文件，填写完成后上传导入。示例文件名：**{EXAMPLE_MAPPING_FILENAME}**")

    st.divider()

    # 导入新配置
    with st.expander("导入映射配置", expanded=True):
        uploaded_file = st.file_uploader(
            "选择 JSON 映射配置文件",
            type=["json"],
            key="mapping_upload"
        )

        if uploaded_file:
            try:
                config_data = json.load(uploaded_file)
                config = MappingConfig.from_dict(config_data)

                name = st.text_input(
                    "配置名称",
                    value=uploaded_file.name.replace(".json", ""),
                    key="mapping_name_input"
                )

                if st.button("保存", key="save_mapping"):
                    item = MappingConfigItem.create(name=name, config=config)
                    store.save_mapping_config(item)
                    st.success(f"映射配置 '{name}' 保存成功")
                    st.rerun()

            except json.JSONDecodeError as e:
                st.error(f"❌ JSON 格式错误: {str(e)}")
            except (KeyError, TypeError, ValueError) as e:
                st.error(f"❌ 配置文件内容错误，缺少必要字段或值无效: {str(e)}")
            except Exception as e:
                st.error(f"❌ 导入失败: {type(e).__name__}: {str(e)}")

    st.divider()

    # 已有的映射配置
    st.subheader("已有映射配置")
    configs = store.load_mapping_configs()

    if not configs:
        st.info("暂无已保存的映射配置")
        return

    # 选择配置
    options = {c.id: c.name for c in configs}
    selected_id = st.selectbox(
        "选择映射配置",
        options=list(options.keys()),
        format_func=lambda x: options[x],
        key="mapping_selector"
    )

    selected_config = next((c for c in configs if c.id == selected_id), None)

    if selected_config:
        # 显示详情
        st.markdown(f"**配置名称:** {selected_config.name}")
        st.markdown(f"**包含映射数:** {len(selected_config.config.sheet_mappings)} 个")

        # 显示每个映射的详细信息
        st.markdown("**映射详情:**")
        for i, mapping in enumerate(selected_config.config.sheet_mappings):
            rule = mapping.copy_rule
            st.markdown(f"- {i+1}. {mapping.data_source_sheet} → {mapping.template_file}")
            st.markdown(f"  规则: {rule.source_start} → {rule.target_start}, {rule.direction.value}, {rule.length}格")

        st.divider()

        # 修改和删除按钮
        col1, col2 = st.columns([1, 1])

        with col1:
            if st.button("修改", key="edit_mapping"):
                st.session_state['edit_mapping_id'] = selected_config.id

        with col2:
            if st.button("删除", key="delete_mapping"):
                store.delete_mapping_config(selected_config.id)
                st.success("删除成功")
                st.rerun()

        # 配置文件 JSON 预览
        with st.expander("查看 JSON 配置"):
            st.json(selected_config.config.to_dict())

    # 修改模式
    if 'edit_mapping_id' in st.session_state:
        edit_id = st.session_state['edit_mapping_id']
        edit_config = store.get_mapping_config(edit_id)

        if edit_config:
            st.divider()
            st.subheader("修改映射配置")

            with st.expander("修改配置", expanded=True):
                # 配置名称
                new_name = st.text_input("配置名称", value=edit_config.name, key="edit_name")

                # 加载数据源和模板，用于下拉选择
                data_sources = store.load_data_sources()
                templates = store.load_templates()

                if not data_sources or not templates:
                    st.warning("请先添加数据源和模板")
                else:
                    # 解析值为数组或字符串的辅助函数
                    def normalize_value(val):
                        if isinstance(val, list):
                            return val
                        elif isinstance(val, str):
                            items = [v.strip() for v in val.split(',')]
                            return items if len(items) > 1 else items[0] if items else val
                        return val

                    # 初始化 session_state 中的映射列表
                    if 'edit_mappings' not in st.session_state:
                        st.session_state['edit_mappings'] = [
                            {
                                "data_source_sheet": m.data_source_sheet,
                                "template_file": m.template_file,
                                "copy_rule": {
                                    "source_start": normalize_value(m.copy_rule.source_start),
                                    "target_start": normalize_value(m.copy_rule.target_start),
                                    "direction": m.copy_rule.direction.value,
                                    "length": m.copy_rule.length
                                }
                            }
                            for m in edit_config.config.sheet_mappings
                        ]

                    # 显示所有映射
                    mappings_to_remove = []
                    for i, mapping in enumerate(st.session_state['edit_mappings']):
                        st.markdown(f"**映射 {i+1}:**")

                        # 数据源选择 - 显示数据源名称和sheet列表
                        ds_options = {ds.id: f"{ds.name} ({len(ds.sheets)} sheets)" for ds in data_sources}
                        # 找到当前映射选中的数据源
                        current_ds_id = None
                        for ds_id, ds_info in ds_options.items():
                            ds_obj = next((d for d in data_sources if d.id == ds_id), None)
                            if ds_obj and mapping.get('data_source_sheet', '') in ds_obj.sheets:
                                current_ds_id = ds_id
                                break
                        if current_ds_id is None and data_sources:
                            current_ds_id = data_sources[0].id

                        selected_ds_id = st.selectbox(
                            f"数据源 {i+1}",
                            options=list(ds_options.keys()),
                            format_func=lambda x: ds_options[x],
                            key=f"edit_ds_{i}",
                            index=list(ds_options.keys()).index(current_ds_id) if current_ds_id in ds_options else 0
                        )

                        # 获取选中的数据源的 sheet 列表
                        selected_ds = next((ds for ds in data_sources if ds.id == selected_ds_id), None)
                        sheet_options = selected_ds.sheets if selected_ds else []

                        # 数据源的 Sheet 选择
                        if sheet_options:
                            # 找到当前映射选中的 sheet
                            current_sheet_idx = 0
                            current_sheet = mapping.get('data_source_sheet', '')
                            if current_sheet in sheet_options:
                                current_sheet_idx = sheet_options.index(current_sheet)

                            selected_sheet = st.selectbox(
                                f"数据源Sheet {i+1}",
                                options=sheet_options,
                                index=current_sheet_idx,
                                key=f"edit_sheet_{i}"
                            )
                        else:
                            selected_sheet = ""

                        # 模板文件选择
                        t_options = {t.id: f"{t.name}" for t in templates}
                        current_t_id = None
                        for t_id, t_name in t_options.items():
                            t_obj = next((tm for tm in templates if tm.id == t_id), None)
                            if t_obj and os.path.basename(t_obj.file_path) == mapping.get('template_file', ''):
                                current_t_id = t_id
                                break
                        if current_t_id is None and templates:
                            current_t_id = templates[0].id

                        selected_t_id = st.selectbox(
                            f"模板文件 {i+1}",
                            options=list(t_options.keys()),
                            format_func=lambda x: t_options[x],
                            key=f"edit_t_{i}",
                            index=list(t_options.keys()).index(current_t_id) if current_t_id in t_options else 0
                        )

                        selected_t = next((t for t in templates if t.id == selected_t_id), None)

                        # 复制规则
                        col_a, col_b, col_c, col_d, col_e = st.columns([3, 3, 2, 2, 1])

                        # 处理 source_start 和 target_start 的值（支持字符串或数组）
                        src_val = mapping.get('copy_rule', {}).get('source_start', 'A1')
                        tgt_val = mapping.get('copy_rule', {}).get('target_start', 'A1')
                        if isinstance(src_val, list):
                            src_val = ", ".join(src_val)
                        if isinstance(tgt_val, list):
                            tgt_val = ", ".join(tgt_val)

                        with col_a:
                            source_start = st.text_input(
                                "源起始（多个用逗号分隔）",
                                value=src_val,
                                key=f"edit_src_{i}",
                                help="例如: A1, B1, C1 或单个值如 A1"
                            )
                        with col_b:
                            target_start = st.text_input(
                                "目标起始（多个用逗号分隔）",
                                value=tgt_val,
                                key=f"edit_tgt_{i}",
                                help="例如: D1, E1, F1 或单个值如 D1"
                            )
                        with col_c:
                            direction = st.selectbox(
                                "方向",
                                options=["horizontal", "vertical"],
                                index=0 if mapping.get('copy_rule', {}).get('direction', 'horizontal') == "horizontal" else 1,
                                key=f"edit_dir_{i}"
                            )
                        with col_d:
                            length = st.number_input("长度", value=mapping.get('copy_rule', {}).get('length', 10), min_value=1, key=f"edit_len_{i}")
                        with col_e:
                            st.markdown("")  # spacing
                            st.markdown("")
                            if st.button("🗑️", key=f"del_mapping_{i}"):
                                mappings_to_remove.append(i)

                        # 更新 session_state
                        st.session_state['edit_mappings'][i] = {
                            "data_source_sheet": selected_sheet if selected_sheet else mapping.get('data_source_sheet', ''),
                            "template_file": os.path.basename(selected_t.file_path) if selected_t else mapping.get('template_file', ''),
                            "copy_rule": {
                                "source_start": normalize_value(source_start),
                                "target_start": normalize_value(target_start),
                                "direction": direction,
                                "length": int(length)
                            }
                        }

                        st.divider()

                    # 添加新映射按钮
                    if st.button("➕ 添加新映射", key="add_mapping"):
                        st.session_state['edit_mappings'].append({
                            "data_source_sheet": data_sources[0].sheets[0] if data_sources and data_sources[0].sheets else "",
                            "template_file": os.path.basename(templates[0].file_path) if templates else "",
                            "copy_rule": {
                                "source_start": ["A1", "B1"],
                                "target_start": ["D1", "E1"],
                                "direction": "horizontal",
                                "length": 10
                            }
                        })
                        st.rerun()

                    # 删除按钮触发后执行删除
                    if mappings_to_remove:
                        for idx in sorted(mappings_to_remove, reverse=True):
                            st.session_state['edit_mappings'].pop(idx)
                        st.rerun()

                    # 保存修改
                    if st.button("💾 保存修改", key="save_edit"):
                        if not st.session_state.get('edit_mappings'):
                            st.error("请至少保留一条映射")
                        else:
                            # 构建新的配置
                            new_config = MappingConfig.from_dict({
                                "version": "1.0",
                                "default_direction": edit_config.config.default_direction.value,
                                "default_write_mode": edit_config.config.default_write_mode.value,
                                "sheet_mappings": st.session_state['edit_mappings']
                            })

                            # 更新配置项
                            updated_item = MappingConfigItem(
                                id=edit_config.id,
                                name=new_name,
                                config=new_config,
                                created_at=edit_config.created_at,
                                updated_at=datetime.now().isoformat()
                            )

                            store.update_mapping_config(updated_item)
                            st.success("修改成功")
                            del st.session_state['edit_mapping_id']
                            if 'edit_mappings' in st.session_state:
                                del st.session_state['edit_mappings']
                            st.rerun()

                    if st.button("取消", key="cancel_edit"):
                        if 'edit_mappings' in st.session_state:
                            del st.session_state['edit_mappings']
                        del st.session_state['edit_mapping_id']
                        st.rerun()
