"""tkinter + ttkbootstrap UI 主应用"""
import os
import sys
import tempfile
import shutil
import json
from datetime import datetime
from typing import Optional

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

from config import STORAGE_DIR, DATA_SOURCES_FILE, TEMPLATES_FILE, MAPPING_CONFIGS_FILE, EXAMPLE_MAPPING_FILENAME, DEFAULT_DIRECTION, DEFAULT_WRITE_MODE
from storage.store import StorageManager
from core.models import DataSourceItem, TemplateItem, MappingConfigItem, MappingConfig, SheetMapping, CopyRule
from core.enums import Direction, WriteMode
from core.copier import DataCopier
from utils.excel_utils import get_sheet_names, load_excel
from utils.path_utils import file_exists, is_file_locked


class ExcelMigratorApp:
    """Excel 数据迁移工具 - tkinter UI"""

    def __init__(self):
        self.store = StorageManager(str(STORAGE_DIR))
        self.current_page = None
        self.temp_files = []  # 跟踪临时文件

        # 创建主窗口
        self.root = ttk.Window(themename="cosmo")
        self.root.title("Excel 数据迁移工具 v1.0")
        self.root.geometry("1000x700")
        self.root.minsize(900, 600)

        self._create_menu()
        self._create_sidebar()
        self._show_page("datasource")

        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        self.root.mainloop()

    def _create_menu(self):
        """创建菜单栏"""
        menubar = ttk.Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = ttk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="退出", command=self._on_closing)

        help_menu = ttk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="关于", command=self._show_about)

    def _create_sidebar(self):
        """创建侧边栏导航"""
        sidebar = ttk.Frame(self.root, bootstyle="secondary")
        sidebar.pack(side=LEFT, fill=Y, padx=0, pady=0)

        # 标题
        title_label = ttk.Label(
            sidebar,
            text="Excel 数据迁移工具",
            font=("Microsoft YaHei", 14, "bold"),
            bootstyle="inverse-secondary"
        )
        title_label.pack(pady=20, padx=10)

        # 导航按钮
        nav_items = [
            ("数据源管理", "datasource", "📊"),
            ("模板管理", "template", "📄"),
            ("映射配置", "mapping", "🔗"),
            ("执行迁移", "migrate", "🚀"),
        ]

        self.nav_buttons = {}
        for text, page_key, icon in nav_items:
            btn = ttk.Button(
                sidebar,
                text=f"{icon} {text}",
                bootstyle="secondary",
                width=15,
                command=lambda p=page_key: self._show_page(p)
            )
            btn.pack(pady=5, padx=10)
            self.nav_buttons[page_key] = btn

        # 版本信息
        version_label = ttk.Label(
            sidebar,
            text="v1.0",
            bootstyle="secondary"
        )
        version_label.pack(side=BOTTOM, pady=10)

    def _show_page(self, page_key: str):
        """切换页面"""
        # 更新按钮状态
        for key, btn in self.nav_buttons.items():
            if key == page_key:
                btn.configure(bootstyle="primary")
            else:
                btn.configure(bootstyle="secondary")

        # 销毁当前页面
        if self.current_page:
            self.current_page.destroy()

        # 创建新页面
        if page_key == "datasource":
            self.current_page = self._create_datasource_page()
        elif page_key == "template":
            self.current_page = self._create_template_page()
        elif page_key == "mapping":
            self.current_page = self._create_mapping_page()
        elif page_key == "migrate":
            self.current_page = self._create_migrate_page()

        self.current_page.pack(side=RIGHT, fill=BOTH, expand=True, padx=10, pady=10)

    def _show_about(self):
        """显示关于对话框"""
        messagebox.showinfo(
            "关于",
            "Excel 数据迁移工具 v1.0\n\n"
            "用于 Excel 数据迁移的工具软件"
        )

    def _on_closing(self):
        """关闭窗口时清理临时文件"""
        if self.temp_files:
            for f in self.temp_files:
                try:
                    if os.path.exists(f):
                        os.remove(f)
                except:
                    pass
        self.root.destroy()

    # ==================== 数据源管理页面 ====================

    def _create_datasource_page(self):
        """创建数据源管理页面"""
        frame = ttk.Frame(self.root)

        # 标题
        ttk.Label(frame, text="数据源管理", font=("Microsoft YaHei", 16, "bold")).pack(anchor=W, pady=10)

        # 配置路径提示
        ttk.Label(frame, text=f"📁 配置文件: {DATA_SOURCES_FILE}", bootstyle="success").pack(anchor=W, pady=5)

        # 上传区域
        upload_frame = ttk.LabelFrame(frame, text="上传数据源")
        upload_frame.pack(fill=X, pady=10)

        file_path_var = tk.StringVar()
        file_name_var = tk.StringVar()

        def select_file():
            path = filedialog.askopenfilename(
                title="选择 Excel 文件",
                filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")]
            )
            if path:
                file_path_var.set(path)
                file_name_var.set(os.path.basename(path).rsplit('.', 1)[0])

        ttk.Button(upload_frame, text="📤 选择文件", command=select_file).pack(side=LEFT)
        ttk.Entry(upload_frame, textvariable=file_path_var, width=60).pack(side=LEFT, padx=10, fill=X, expand=True)

        # 名称输入
        name_frame = ttk.Frame(frame)
        name_frame.pack(fill=X, pady=10)
        ttk.Label(name_frame, text="数据源名称:").pack(side=LEFT)
        name_entry = ttk.Entry(name_frame, textvariable=file_name_var, width=30)
        name_entry.pack(side=LEFT, padx=10)

        # 按钮
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=X, pady=10)
        save_btn = ttk.Button(btn_frame, text="💾 保存", bootstyle="success", command=lambda: self._save_datasource(file_path_var.get(), file_name_var.get(), upload_frame, ds_list_frame))
        save_btn.pack(side=LEFT, padx=5)

        # 数据源列表
        ds_list_frame = ttk.LabelFrame(frame, text="已保存的数据源")
        ds_list_frame.pack(fill=BOTH, expand=True, pady=10)

        self._refresh_datasource_list(ds_list_frame)

        return frame

    def _refresh_datasource_list(self, parent_frame):
        """刷新数据源列表"""
        # 清除现有内容
        for widget in parent_frame.winfo_children():
            widget.destroy()

        data_sources = self.store.load_data_sources()

        if not data_sources:
            ttk.Label(parent_frame, text="暂无已保存的数据源", bootstyle="info").pack()
            return

        # 创建表格
        columns = ("name", "path", "sheets", "status")
        tree = ttk.Treeview(parent_frame, columns=columns, show="headings", height=10)

        tree.heading("name", text="名称")
        tree.heading("path", text="路径")
        tree.heading("sheets", text="Sheet数量")
        tree.heading("status", text="状态")

        tree.column("name", width=150)
        tree.column("path", width=300)
        tree.column("sheets", width=80, anchor=CENTER)
        tree.column("status", width=100, anchor=CENTER)

        for ds in data_sources:
            # 检查文件状态
            status = "✅ 正常"
            if not file_exists(ds.file_path):
                status = "⚠️ 文件不存在"
            elif is_file_locked(ds.file_path):
                status = "⚠️ 文件被占用"

            tree.insert("", END, values=(ds.name, ds.file_path, len(ds.sheets), status), tags=(ds.id,))

        tree.pack(fill=BOTH, expand=True)

        # 按钮
        btn_frame = ttk.Frame(parent_frame)
        btn_frame.pack(fill=X, pady=5)

        def on_update():
            selected = tree.selection()
            if selected:
                ds_id = tree.item(selected[0])["tags"][0]
                self._update_datasource(ds_id, parent_frame)

        def on_delete():
            selected = tree.selection()
            if selected:
                ds_id = tree.item(selected[0])["tags"][0]
                if messagebox.askyesno("确认", "确定要删除这个数据源吗？"):
                    self.store.delete_data_source(ds_id)
                    self._refresh_datasource_list(parent_frame)

        ttk.Button(btn_frame, text="🔄 更新", command=on_update).pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="🗑️ 删除", bootstyle="danger", command=on_delete).pack(side=LEFT)

    def _save_datasource(self, file_path: str, name: str, upload_frame, list_frame):
        """保存数据源"""
        if not file_path:
            messagebox.showerror("错误", "请选择文件")
            return
        if not name:
            messagebox.showerror("错误", "请输入名称")
            return

        if not file_exists(file_path):
            messagebox.showerror("错误", "文件不存在")
            return
        if is_file_locked(file_path):
            messagebox.showerror("错误", "文件已被打开，请关闭后重试")
            return

        try:
            wb = load_excel(file_path)
            sheets = get_sheet_names(wb)
            del wb

            # 复制到临时目录
            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, os.path.basename(file_path))
            shutil.copy2(file_path, temp_path)
            self.temp_files.append(temp_path)

            item = DataSourceItem.create(name=name, file_path=temp_path, sheets=sheets)
            is_new = self.store.save_data_source(item)

            if is_new:
                messagebox.showinfo("成功", f"新增数据源 '{name}'，包含 {len(sheets)} 个 Sheet")
            else:
                messagebox.showinfo("成功", f"更新数据源 '{name}'，包含 {len(sheets)} 个 Sheet")

            self._refresh_datasource_list(list_frame)

        except Exception as e:
            messagebox.showerror("错误", f"处理失败: {str(e)}")

    def _update_datasource(self, ds_id: str, parent_frame):
        """更新数据源"""
        ds = self.store.get_data_source(ds_id)
        if not ds:
            return

        if not file_exists(ds.file_path):
            messagebox.showerror("错误", "文件不存在，无法更新")
            return
        if is_file_locked(ds.file_path):
            messagebox.showerror("错误", "文件已打开，请关闭后重试")
            return

        try:
            wb = load_excel(ds.file_path)
            actual_sheets = get_sheet_names(wb)
            del wb

            from core.models import DataSourceItem as DSItem
            updated_item = DSItem(
                id=ds.id,
                name=ds.name,
                file_path=ds.file_path,
                sheets=actual_sheets,
                added_at=ds.added_at
            )
            self.store.update_data_source(updated_item)
            messagebox.showinfo("成功", f"已更新，当前共 {len(actual_sheets)} 个 Sheet")
            self._refresh_datasource_list(parent_frame)
        except Exception as e:
            messagebox.showerror("错误", f"更新失败: {str(e)}")

    # ==================== 模板管理页面 ====================

    def _create_template_page(self):
        """创建模板管理页面"""
        frame = ttk.Frame(self.root)

        ttk.Label(frame, text="模板管理", font=("Microsoft YaHei", 16, "bold")).pack(anchor=W, pady=10)
        ttk.Label(frame, text=f"📁 配置文件: {TEMPLATES_FILE}", bootstyle="success").pack(anchor=W, pady=5)

        upload_frame = ttk.LabelFrame(frame, text="上传模板")
        upload_frame.pack(fill=X, pady=10)

        file_path_var = tk.StringVar()
        file_name_var = tk.StringVar()

        def select_file():
            path = filedialog.askopenfilename(
                title="选择 Excel 文件",
                filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")]
            )
            if path:
                file_path_var.set(path)
                file_name_var.set(os.path.basename(path).rsplit('.', 1)[0])

        ttk.Button(upload_frame, text="📤 选择文件", command=select_file).pack(side=LEFT)
        ttk.Entry(upload_frame, textvariable=file_path_var, width=60).pack(side=LEFT, padx=10, fill=X, expand=True)

        name_frame = ttk.Frame(frame)
        name_frame.pack(fill=X, pady=10)
        ttk.Label(name_frame, text="模板名称:").pack(side=LEFT)
        name_entry = ttk.Entry(name_frame, textvariable=file_name_var, width=30)
        name_entry.pack(side=LEFT, padx=10)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=X, pady=10)
        ttk.Button(btn_frame, text="💾 保存", bootstyle="success", command=lambda: self._save_template(file_path_var.get(), file_name_var.get(), tmpl_list_frame)).pack(side=LEFT, padx=5)

        tmpl_list_frame = ttk.LabelFrame(frame, text="已保存的模板")
        tmpl_list_frame.pack(fill=BOTH, expand=True, pady=10)

        self._refresh_template_list(tmpl_list_frame)

        return frame

    def _refresh_template_list(self, parent_frame):
        for widget in parent_frame.winfo_children():
            widget.destroy()

        templates = self.store.load_templates()

        if not templates:
            ttk.Label(parent_frame, text="暂无已保存的模板", bootstyle="info").pack()
            return

        columns = ("name", "path", "sheets", "status")
        tree = ttk.Treeview(parent_frame, columns=columns, show="headings", height=10)

        tree.heading("name", text="名称")
        tree.heading("path", text="路径")
        tree.heading("sheets", text="Sheet数量")
        tree.heading("status", text="状态")

        tree.column("name", width=150)
        tree.column("path", width=300)
        tree.column("sheets", width=80, anchor=CENTER)
        tree.column("status", width=100, anchor=CENTER)

        for t in templates:
            status = "✅ 正常"
            if not file_exists(t.file_path):
                status = "⚠️ 文件不存在"
            elif is_file_locked(t.file_path):
                status = "⚠️ 文件被占用"

            tree.insert("", END, values=(t.name, t.file_path, len(t.sheets), status), tags=(t.id,))

        tree.pack(fill=BOTH, expand=True)

        btn_frame = ttk.Frame(parent_frame)
        btn_frame.pack(fill=X, pady=5)

        def on_update():
            selected = tree.selection()
            if selected:
                t_id = tree.item(selected[0])["tags"][0]
                self._update_template(t_id, parent_frame)

        def on_delete():
            selected = tree.selection()
            if selected:
                t_id = tree.item(selected[0])["tags"][0]
                if messagebox.askyesno("确认", "确定要删除这个模板吗？"):
                    self.store.delete_template(t_id)
                    self._refresh_template_list(parent_frame)

        ttk.Button(btn_frame, text="🔄 更新", command=on_update).pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="🗑️ 删除", bootstyle="danger", command=on_delete).pack(side=LEFT)

    def _save_template(self, file_path: str, name: str, list_frame):
        if not file_path or not name:
            messagebox.showerror("错误", "请选择文件并输入名称")
            return

        if not file_exists(file_path) or is_file_locked(file_path):
            messagebox.showerror("错误", "文件不存在或已被打开")
            return

        try:
            wb = load_excel(file_path)
            sheets = get_sheet_names(wb)
            del wb

            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, os.path.basename(file_path))
            shutil.copy2(file_path, temp_path)
            self.temp_files.append(temp_path)

            item = TemplateItem.create(name=name, file_path=temp_path, sheets=sheets)
            is_new = self.store.save_template(item)

            messagebox.showinfo("成功", f"{'新增' if is_new else '更新'}模板 '{name}'，包含 {len(sheets)} 个 Sheet")
            self._refresh_template_list(list_frame)

        except Exception as e:
            messagebox.showerror("错误", f"处理失败: {str(e)}")

    def _update_template(self, t_id: str, parent_frame):
        t = self.store.get_template(t_id)
        if not t:
            return

        if not file_exists(t.file_path):
            messagebox.showerror("错误", "文件不存在")
            return
        if is_file_locked(t.file_path):
            messagebox.showerror("错误", "文件已打开")
            return

        try:
            wb = load_excel(t.file_path)
            actual_sheets = get_sheet_names(wb)
            del wb

            updated_item = TemplateItem(
                id=t.id, name=t.name, file_path=t.file_path,
                sheets=actual_sheets, added_at=t.added_at
            )
            self.store.update_template(updated_item)
            messagebox.showinfo("成功", f"已更新，当前 {len(actual_sheets)} 个 Sheet")
            self._refresh_template_list(parent_frame)
        except Exception as e:
            messagebox.showerror("错误", f"更新失败: {str(e)}")

    # ==================== 映射配置页面 ====================

    def _create_mapping_page(self):
        """创建映射配置页面"""
        frame = ttk.Frame(self.root)

        ttk.Label(frame, text="映射配置", font=("Microsoft YaHei", 16, "bold")).pack(anchor=W, pady=10)

        # 下载示例
        example_frame = ttk.Frame(frame)
        example_frame.pack(fill=X, pady=5)

        def download_example():
            example_config = {
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
            path = filedialog.asksaveasfilename(
                title="保存示例文件",
                defaultextension=".json",
                initialfile=EXAMPLE_MAPPING_FILENAME,
                filetypes=[("JSON 文件", "*.json")]
            )
            if path:
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(example_config, f, ensure_ascii=False, indent=2)
                messagebox.showinfo("成功", f"示例文件已保存到:\n{path}")

        ttk.Button(example_frame, text="📥 下载示例文件", command=download_example).pack(side=LEFT)
        ttk.Label(example_frame, text=f"请下载示例文件，填写完成后上传导入。示例文件名：{EXAMPLE_MAPPING_FILENAME}", bootstyle="info").pack(side=LEFT, padx=10)

        # 导入新配置
        import_frame = ttk.LabelFrame(frame, text="导入映射配置")
        import_frame.pack(fill=X, pady=10)

        config_path_var = tk.StringVar()
        config_name_var = tk.StringVar()

        def select_config():
            path = filedialog.askopenfilename(
                title="选择映射配置文件",
                filetypes=[("JSON 文件", "*.json")]
            )
            if path:
                config_path_var.set(path)
                config_name_var.set(os.path.basename(path).replace(".json", ""))

        ttk.Button(import_frame, text="📤 选择JSON文件", command=select_config).pack(side=LEFT)
        ttk.Entry(import_frame, textvariable=config_path_var, width=40).pack(side=LEFT, padx=10)
        ttk.Label(import_frame, text="配置名称:").pack(side=LEFT)
        ttk.Entry(import_frame, textvariable=config_name_var, width=20).pack(side=LEFT, padx=10)
        ttk.Button(import_frame, text="💾 保存", bootstyle="success", command=self._import_mapping_config).pack(side=LEFT, padx=10)

        # 已有配置
        list_frame = ttk.LabelFrame(frame, text="已有映射配置")
        list_frame.pack(fill=BOTH, expand=True, pady=10)

        self._refresh_mapping_list(list_frame)

        return frame

    def _import_mapping_config(self):
        config_path = None  # 获取导入文件路径
        name = None

        # 弹出文件选择
        path = filedialog.askopenfilename(title="选择映射配置文件", filetypes=[("JSON 文件", "*.json")])
        if not path:
            return

        # 弹出名称输入对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("导入映射配置")
        dialog.geometry("300x100")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="配置名称:").pack(pady=5)
        name_var = tk.StringVar(value=os.path.basename(path).replace(".json", ""))
        ttk.Entry(dialog, textvariable=name_var, width=30).pack(pady=5)

        def do_import():
            config_name = name_var.get().strip()
            if not config_name:
                messagebox.showerror("错误", "请输入配置名称")
                return

            try:
                with open(path, "r", encoding="utf-8") as f:
                    config_data = json.load(f)

                config = MappingConfig.from_dict(config_data)
                item = MappingConfigItem.create(name=config_name, config=config)
                self.store.save_mapping_config(item)
                messagebox.showinfo("成功", f"映射配置 '{config_name}' 保存成功")
                dialog.destroy()

            except json.JSONDecodeError as e:
                messagebox.showerror("错误", f"JSON 格式错误: {str(e)}")
            except Exception as e:
                messagebox.showerror("错误", f"导入失败: {str(e)}")

        ttk.Button(dialog, text="确定", command=do_import).pack(pady=10)

    def _refresh_mapping_list(self, parent_frame):
        for widget in parent_frame.winfo_children():
            widget.destroy()

        configs = self.store.load_mapping_configs()

        if not configs:
            ttk.Label(parent_frame, text="暂无已保存的映射配置", bootstyle="info").pack()
            return

        # 下拉选择
        select_frame = ttk.Frame(parent_frame)
        select_frame.pack(fill=X, pady=5)

        ttk.Label(select_frame, text="选择映射配置:").pack(side=LEFT)
        selected_id = [None]
        combo = ttk.Combobox(select_frame, values=[c.name for c in configs], state="readonly", width=30)
        combo.current(0)
        combo.pack(side=LEFT, padx=10)
        selected_id[0] = configs[0].id

        def on_select(event):
            idx = combo.current()
            selected_id[0] = configs[idx].id
            self._show_mapping_detail(detail_frame, configs[idx])

        combo.bind("<<ComboboxSelected>>", on_select)

        # 详情区域
        detail_frame = ttk.Frame(parent_frame)
        detail_frame.pack(fill=BOTH, expand=True, pady=10)

        if configs:
            self._show_mapping_detail(detail_frame, configs[0])

    def _show_mapping_detail(self, parent_frame, config_item):
        for widget in parent_frame.winfo_children():
            widget.destroy()

        ttk.Label(parent_frame, text=f"配置名称: {config_item.name}", font=("Microsoft YaHei", 12, "bold")).pack(anchor=W, pady=5)
        ttk.Label(parent_frame, text=f"包含映射数: {len(config_item.config.sheet_mappings)} 个", bootstyle="info").pack(anchor=W)

        ttk.Label(parent_frame, text="映射详情:", font=("Microsoft YaHei", 10, "bold")).pack(anchor=W, pady=10)

        for i, mapping in enumerate(config_item.config.sheet_mappings):
            rule = mapping.copy_rule
            if rule.is_array_mode():
                pairs = rule.get_pairs()
                text = f"{i+1}. {mapping.data_source_sheet} → {mapping.template_file} ({rule.direction.value}, {rule.length}格)"
            else:
                text = f"{i+1}. {mapping.data_source_sheet} → {mapping.template_file} ({rule.source_start} → {rule.target_start}, {rule.direction.value}, {rule.length}格)"
            ttk.Label(parent_frame, text=text).pack(anchor=W, padx=20)

        # JSON 预览
        json_frame = ttk.LabelFrame(parent_frame, text="JSON 配置预览")
        json_frame.pack(fill=BOTH, expand=True, pady=10)

        json_text = scrolledtext.ScrolledText(json_frame, height=10, font=("Consolas", 9))
        json_text.pack(fill=BOTH, expand=True)
        json_text.insert("1.0", json.dumps(config_item.config.to_dict(), ensure_ascii=False, indent=2))
        json_text.config(state="disabled")

        # 按钮
        btn_frame = ttk.Frame(parent_frame)
        btn_frame.pack(fill=X, pady=10)

        ttk.Button(btn_frame, text="✏️ 修改", command=lambda: self._edit_mapping_config(config_item)).pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="🗑️ 删除", bootstyle="danger", command=lambda: self._delete_mapping_config(config_item.id)).pack(side=LEFT)

    def _edit_mapping_config(self, config_item):
        """编辑映射配置"""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"修改映射配置 - {config_item.name}")
        dialog.geometry("800x500")
        dialog.transient(self.root)

        main_frame = ttk.Frame(dialog)
        main_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)

        # 配置名称
        name_frame = ttk.Frame(main_frame)
        name_frame.pack(fill=X, pady=5)
        ttk.Label(name_frame, text="配置名称:").pack(side=LEFT)
        name_var = tk.StringVar(value=config_item.name)
        ttk.Entry(name_frame, textvariable=name_var, width=30).pack(side=LEFT, padx=10)

        # 数据源和模板下拉
        data_sources = self.store.load_data_sources()
        templates = self.store.load_templates()

        mapping_vars = []

        # 创建映射列表容器
        list_canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=list_canvas.yview)
        scrollable_frame = ttk.Frame(list_canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: list_canvas.configure(scrollregion=list_canvas.bbox("all"))
        )

        list_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        list_canvas.configure(yscrollcommand=scrollbar.set)

        # 解析现有映射
        for i, mapping in enumerate(config_item.config.sheet_mappings):
            mapping_frame = ttk.LabelFrame(scrollable_frame, text=f"映射 {i+1}")
            mapping_frame.pack(fill=X, pady=5, padx=5)

            mv = {
                "ds_sheet": tk.StringVar(value=mapping.data_source_sheet),
                "template": tk.StringVar(value=mapping.template_file),
                "source_start": tk.StringVar(value=", ".join(mapping.copy_rule.source_start) if isinstance(mapping.copy_rule.source_start, list) else mapping.copy_rule.source_start),
                "target_start": tk.StringVar(value=", ".join(mapping.copy_rule.target_start) if isinstance(mapping.copy_rule.target_start, list) else mapping.copy_rule.target_start),
                "direction": tk.StringVar(value=mapping.copy_rule.direction.value),
                "length": tk.IntVar(value=mapping.copy_rule.length),
                "index": i
            }
            mapping_vars.append(mv)

            row = 0
            ttk.Label(mapping_frame, text="数据源Sheet:").grid(row=row, column=0, sticky=W, padx=5, pady=2)
            ds_combo = ttk.Combobox(mapping_frame, values=[s for ds in data_sources for s in ds.sheets],
                                   textvariable=mv["ds_sheet"], width=20, state="readonly")
            ds_combo.grid(row=row, column=1, sticky=W, padx=5, pady=2)

            row += 1
            ttk.Label(mapping_frame, text="模板文件:").grid(row=row, column=0, sticky=W, padx=5, pady=2)
            t_combo = ttk.Combobox(mapping_frame, values=[os.path.basename(t.file_path) for t in templates],
                                  textvariable=mv["template"], width=20, state="readonly")
            t_combo.grid(row=row, column=1, sticky=W, padx=5, pady=2)

            row += 1
            ttk.Label(mapping_frame, text="源起始:").grid(row=row, column=0, sticky=W, padx=5, pady=2)
            ttk.Entry(mapping_frame, textvariable=mv["source_start"], width=25).grid(row=row, column=1, sticky=W, padx=5, pady=2)

            row += 1
            ttk.Label(mapping_frame, text="目标起始:").grid(row=row, column=0, sticky=W, padx=5, pady=2)
            ttk.Entry(mapping_frame, textvariable=mv["target_start"], width=25).grid(row=row, column=1, sticky=W, padx=5, pady=2)

            row += 1
            ttk.Label(mapping_frame, text="方向:").grid(row=row, column=0, sticky=W, padx=5, pady=2)
            ttk.Combobox(mapping_frame, values=["horizontal", "vertical"], textvariable=mv["direction"],
                        width=18, state="readonly").grid(row=row, column=1, sticky=W, padx=5, pady=2)

            row += 1
            ttk.Label(mapping_frame, text="长度:").grid(row=row, column=0, sticky=W, padx=5, pady=2)
            ttk.Entry(mapping_frame, textvariable=mv["length"], width=25).grid(row=row, column=1, sticky=W, padx=5, pady=2)

        list_canvas.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        # 按钮
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill=X, pady=10)

        def do_save():
            # 构建新配置
            sheet_mappings = []
            for mv in mapping_vars:
                src = mv["source_start"].get()
                tgt = mv["target_start"].get()
                sheet_mappings.append({
                    "data_source_sheet": mv["ds_sheet"].get(),
                    "template_file": mv["template"].get(),
                    "copy_rule": {
                        "source_start": [x.strip() for x in src.split(",")] if "," in src else src,
                        "target_start": [x.strip() for x in tgt.split(",")] if "," in tgt else tgt,
                        "direction": mv["direction"].get(),
                        "length": mv["length"].get()
                    }
                })

            new_config = MappingConfig.from_dict({
                "version": "1.0",
                "default_direction": config_item.config.default_direction.value,
                "default_write_mode": config_item.config.default_write_mode.value,
                "sheet_mappings": sheet_mappings
            })

            updated_item = MappingConfigItem(
                id=config_item.id,
                name=name_var.get(),
                config=new_config,
                created_at=config_item.created_at,
                updated_at=datetime.now().isoformat()
            )

            self.store.update_mapping_config(updated_item)
            messagebox.showinfo("成功", "修改成功")
            dialog.destroy()

        ttk.Button(btn_frame, text="💾 保存修改", bootstyle="success", command=do_save).pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=LEFT)

    def _delete_mapping_config(self, config_id):
        if messagebox.askyesno("确认", "确定要删除这个映射配置吗？"):
            self.store.delete_mapping_config(config_id)
            messagebox.showinfo("成功", "删除成功")

    # ==================== 执行迁移页面 ====================

    def _create_migrate_page(self):
        """创建执行迁移页面"""
        frame = ttk.Frame(self.root)

        ttk.Label(frame, text="执行迁移", font=("Microsoft YaHei", 16, "bold")).pack(anchor=W, pady=10)

        data_sources = self.store.load_data_sources()
        templates = self.store.load_templates()
        configs = self.store.load_mapping_configs()

        if not data_sources:
            ttk.Label(frame, text="暂无数据源，请先在【数据源管理】中添加", bootstyle="danger").pack(pady=20)
            return frame

        if not templates:
            ttk.Label(frame, text="暂无模板，请先在【模板管理】中添加", bootstyle="danger").pack(pady=20)
            return frame

        if not configs:
            ttk.Label(frame, text="暂无映射配置，请先在【映射配置】中添加", bootstyle="danger").pack(pady=20)
            return frame

        # 选择数据源
        ds_frame = ttk.LabelFrame(frame, text="选择数据源")
        ds_frame.pack(fill=X, pady=10)

        ds_var = tk.StringVar(value=data_sources[0].id)
        ds_combo = ttk.Combobox(ds_frame, values=[(ds.id, ds.name) for ds in data_sources],
                                textvariable=ds_var, state="readonly", width=30)
        ds_combo.pack(side=LEFT)

        # 选择模板
        t_frame = ttk.LabelFrame(frame, text="选择模板（将被修改）")
        t_frame.pack(fill=X, pady=10)

        t_var = tk.StringVar(value=templates[0].id)
        t_combo = ttk.Combobox(t_frame, values=[(t.id, t.name) for t in templates],
                               textvariable=t_var, state="readonly", width=30)
        t_combo.pack(side=LEFT)

        # 选择映射配置
        config_frame = ttk.LabelFrame(frame, text="选择映射配置")
        config_frame.pack(fill=X, pady=10)

        config_var = tk.StringVar(value=configs[0].id)
        config_combo = ttk.Combobox(config_frame, values=[(c.id, c.name) for c in configs],
                                    textvariable=config_var, state="readonly", width=30)
        config_combo.pack(side=LEFT)

        # 映射预览
        preview_frame = ttk.LabelFrame(frame, text="映射预览")
        preview_frame.pack(fill=BOTH, expand=True, pady=10)

        preview_text = scrolledtext.ScrolledText(preview_frame, height=15, font=("Consolas", 9))
        preview_text.pack(fill=BOTH, expand=True)

        def update_preview():
            preview_text.config(state="normal")
            preview_text.delete("1.0", END)

            ds_id = ds_var.get()
            config_id = config_var.get()

            ds = self.store.get_data_source(ds_id)
            config_item = self.store.get_mapping_config(config_id)

            if not ds or not config_item:
                preview_text.insert("1.0", "请选择数据源和映射配置")
                preview_text.config(state="disabled")
                return

            template_name_map = {os.path.basename(t.file_path): t for t in templates}

            matched = []
            unmatched_src = []
            unmatched_tgt = []

            for mapping in config_item.config.sheet_mappings:
                if mapping.data_source_sheet not in ds.sheets:
                    unmatched_src.append(mapping.data_source_sheet)
                    continue
                if mapping.template_file not in template_name_map:
                    unmatched_tgt.append(mapping.template_file)
                    continue
                matched.append(mapping)

            if matched:
                preview_text.insert(END, f"✅ 将迁移 {len(matched)} 个 Sheet:\n\n")
                for m in matched:
                    rule = m.copy_rule
                    preview_text.insert(END, f"• {m.data_source_sheet} → {m.template_file}\n")
                    preview_text.insert(END, f"  {rule.source_start} → {rule.target_start}, {rule.direction.value}, {rule.length}格\n\n")

            if unmatched_src:
                preview_text.insert(END, f"\n⚠️ 数据源中不存在 {len(unmatched_src)} 个 Sheet:\n")
                for s in unmatched_src:
                    preview_text.insert(END, f"  - {s}\n")

            if unmatched_tgt:
                preview_text.insert(END, f"\n⚠️ 未找到对应模板文件 {len(unmatched_tgt)} 个:\n")
                for f in unmatched_tgt:
                    preview_text.insert(END, f"  - {f}\n")

            preview_text.config(state="disabled")

        ds_combo.bind("<<ComboboxSelected>>", lambda e: update_preview())
        config_combo.bind("<<ComboboxSelected>>", lambda e: update_preview())

        # 执行按钮
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=X, pady=10)

        result_text = scrolledtext.ScrolledText(frame, height=8, font=("Consolas", 9), state="disabled")
        result_text.pack(fill=X, pady=10)

        def execute_migrate():
            result_text.config(state="normal")
            result_text.delete("1.0", END)
            result_text.update()

            ds_id = ds_var.get()
            t_id = t_var.get()
            config_id = config_var.get()

            ds = self.store.get_data_source(ds_id)
            t_item = self.store.get_template(t_id)
            config_item = self.store.get_mapping_config(config_id)

            if not all([ds, t_item, config_item]):
                result_text.insert("1.0", "请确保已选择数据源、模板和映射配置")
                result_text.config(state="disabled")
                return

            if not file_exists(ds.file_path):
                result_text.insert("1.0", f"❌ 数据源文件不存在: {ds.file_path}")
                result_text.config(state="disabled")
                return

            if not file_exists(t_item.file_path):
                result_text.insert("1.0", f"❌ 模板文件不存在: {t_item.file_path}")
                result_text.config(state="disabled")
                return

            if is_file_locked(ds.file_path):
                result_text.insert("1.0", "❌ 数据源文件已打开，请关闭后重试")
                result_text.config(state="disabled")
                return

            try:
                source_wb = load_excel(ds.file_path, data_only=True)

                template_name_map = {os.path.basename(t.file_path): t for t in templates}
                target_path = t_item.file_path

                # 如果模板被占用，复制一份
                copied = False
                if is_file_locked(target_path):
                    copy_path = _generate_copy_path(target_path)
                    shutil.copy2(target_path, copy_path)
                    target_path = copy_path
                    copied = True
                    result_text.insert(END, f"📁 原始文件被占用，已生成副本: {copy_path}\n\n")
                    result_text.update()

                target_wb = load_excel(target_path)

                copier = DataCopier(write_mode=WriteMode.OVERWRITE)
                all_results = []

                template_groups = {}
                for mapping in config_item.config.sheet_mappings:
                    t_file = mapping.template_file
                    if t_file not in template_groups:
                        template_groups[t_file] = []
                    template_groups[t_file].append(mapping)

                for t_file, mappings in template_groups.items():
                    if t_file not in template_name_map:
                        continue

                    for mapping in mappings:
                        if mapping.data_source_sheet not in ds.sheets:
                            continue
                        result = copier.copy(source_wb, target_wb, mapping)
                        all_results.append(result)

                target_wb.save(target_path)

                # 显示结果
                success = [r for r in all_results if r.success]
                failed = [r for r in all_results if not r.success]

                result_text.insert(END, "=" * 40 + "\n")
                if success:
                    result_text.insert(END, f"✅ 成功: {len(success)} 个 Sheet\n")
                    for r in success:
                        result_text.insert(END, f"  - {r.source_sheet} → {r.template_file}\n")

                if failed:
                    result_text.insert(END, f"\n⚠️ 跳过: {len(failed)} 个 Sheet\n")
                    for r in failed:
                        result_text.insert(END, f"  - {r.source_sheet}: {r.message}\n")

                result_text.insert(END, f"\n{'=' * 40}\n")
                result_text.insert(END, f"总计: {len(all_results)} | 成功: {len(success)} | 跳过: {len(failed)}\n")

            except PermissionError:
                result_text.insert("1.0", "⚠️ 文件已打开，请关闭后重试")
            except Exception as e:
                result_text.insert("1.0", f"❌ 迁移失败: {str(e)}")

            result_text.config(state="disabled")

        ttk.Button(btn_frame, text="🚀 执行迁移", bootstyle="success", command=execute_migrate).pack(side=LEFT, padx=5)

        update_preview()

        return frame


def _generate_copy_path(original_path: str) -> str:
    """生成副本文件路径"""
    directory = os.path.dirname(original_path)
    filename = os.path.basename(original_path)
    name, ext = os.path.splitext(filename)
    copy_filename = f"{name}_副本{ext}"
    return os.path.join(directory, copy_filename)


def main():
    """主入口"""
    ExcelMigratorApp()


if __name__ == "__main__":
    main()
