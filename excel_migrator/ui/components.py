"""UI 公共组件"""
import os
import streamlit as st

from core.models import DataSourceItem, TemplateItem, MappingConfigItem
from utils.path_utils import file_exists, is_file_locked


def try_load_excel(file_path: str, data_only: bool = False) -> tuple:
    """
    尝试加载 Excel 文件
    返回 (workbook, error_message)
    data_only=True 时读取公式的计算值（而不是公式本身）
    """
    if not file_path:
        return None, "请选择文件"

    if not file_exists(file_path):
        return None, "文件不存在"

    if is_file_locked(file_path):
        return None, "文件已打开，请关闭后重试"

    try:
        from utils.excel_utils import load_excel
        wb = load_excel(file_path, data_only=data_only)
        return wb, None
    except PermissionError:
        return None, "文件已打开，请关闭后重试"
    except Exception as e:
        return None, f"读取失败: {str(e)}"


def render_file_upload_simple(
    label: str,
    key: str
) -> tuple:
    """
    简单的文件上传组件
    返回: (file_path, file_name, error)
    """
    st.info(f"📤 上传 {label}（文件会被复制到临时目录）")

    uploaded_file = st.file_uploader(
        f"选择 {label} 文件",
        type=["xlsx", "xlsm"],
        key=f"{key}_upload",
        help="支持的格式: .xlsx, .xlsm"
    )

    if uploaded_file is None:
        return None, None, "请选择文件"

    # 保存到临时目录
    import tempfile
    temp_dir = tempfile.gettempdir()
    file_path = os.path.join(temp_dir, uploaded_file.name)

    # 检查临时文件是否被占用
    if is_file_locked(file_path):
        return None, None, "文件已打开，请关闭后重试"

    try:
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # 验证文件可以读取
        wb, error = try_load_excel(file_path)
        if error:
            return None, None, f"文件无效: {error}"

        del wb
        file_name = uploaded_file.name.replace(".xlsx", "").replace(".xlsm", "")

        return file_path, file_name, None

    except Exception as e:
        return None, None, f"上传失败: {str(e)}"


def render_result_message(results: list) -> None:
    """渲染迁移结果消息"""
    if not results:
        st.info("没有可显示的结果")
        return

    success_results = [r for r in results if r.success]
    failed_results = [r for r in results if not r.success]

    if success_results:
        st.success(f"✅ 成功: {len(success_results)} 个 Sheet")
        for r in success_results:
            st.markdown(f"- {r.source_sheet} → {r.template_file}")

    if failed_results:
        st.warning(f"⚠️ 跳过: {len(failed_results)} 个 Sheet")
        for r in failed_results:
            st.markdown(f"- {r.source_sheet}: {r.message}")

    st.divider()
    st.markdown(f"**总计:** {len(results)} | 成功: {len(success_results)} | 跳过: {len(failed_results)}")
