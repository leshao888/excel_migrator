"""tkinter + ttkbootstrap UI 主应用"""
import os
import sys
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
from core.enums import Direction, WriteMode, normalize_string, normalize_list_string
from core.copier import DataCopier
from utils.excel_utils import get_sheet_names, load_excel
from utils.path_utils import file_exists, is_file_locked


class ExcelMigratorApp:
    """Excel 数据迁移工具 - tkinter UI"""

    def __init__(self):
        self.store = StorageManager(str(STORAGE_DIR))
        self.current_page = None

        # 创建主窗口
        self.root = ttk.Window(themename="cosmo")
        self.root.title("Excel 数据迁移工具 v1.1")
        self.root.geometry("1100x750")
        self.root.minsize(900, 600)

        self._create_menu()
        self._create_sidebar()
        self._show_page("datasource")

        self.root.mainloop()

    def _create_menu(self):
        """创建菜单栏"""
        menubar = ttk.Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = ttk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="退出", command=self.root.destroy)

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
        version_label = ttk.Label(sidebar, text="v1.1", bootstyle="secondary")
        version_label.pack(side=BOTTOM, pady=10)

    def _show_page(self, page_key: str):
        """切换页面"""
        for key, btn in self.nav_buttons.items():
            btn.configure(bootstyle="primary" if key == page_key else "secondary")

        if self.current_page:
            self.current_page.destroy()

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
        messagebox.showinfo("关于", "Excel 数据迁移工具 v1.1\n\n用于 Excel 数据迁移的工具软件")

    # ==================== 数据源管理页面 ====================

    def _create_datasource_page(self):
        """创建数据源管理页面"""
        frame = ttk.Frame(self.root)

        ttk.Label(frame, text="数据源管理", font=("Microsoft YaHei", 16, "bold")).pack(anchor=W, pady=10)
        ttk.Label(frame, text="添加 Excel 文件作为数据源，支持单文件或多文件导入", bootstyle="info").pack(anchor=W, pady=5)

        # 导入区域
        import_frame = ttk.LabelFrame(frame, text="导入数据源")
        import_frame.pack(fill=X, pady=10)

        btn_import_single = ttk.Button(import_frame, text="📂 导入单文件", bootstyle="success",
                                       command=lambda: self._import_datasource_files(multi=False))
        btn_import_single.pack(side=LEFT, padx=5, pady=10)

        btn_import_multi = ttk.Button(import_frame, text="📂 导入多文件", bootstyle="success",
                                       command=lambda: self._import_datasource_files(multi=True))
        btn_import_multi.pack(side=LEFT, padx=5, pady=10)

        # 列表区域
        list_frame = ttk.LabelFrame(frame, text="已保存的数据源")
        list_frame.pack(fill=BOTH, expand=True, pady=10)

        self._refresh_datasource_list(list_frame)

        return frame

    def _import_datasource_files(self, multi=False):
        """导入数据源文件"""
        if multi:
            paths = filedialog.askopenfilenames(title="选择数据源文件", filetypes=[("Excel 文件", "*.xlsx *.xlsm")])
        else:
            path = filedialog.askopenfilename(title="选择数据源文件", filetypes=[("Excel 文件", "*.xlsx *.xlsm")])
            paths = [path] if path else []

        if not paths:
            return

        success_count = 0
        for file_path in paths:
            if not file_path:
                continue

            if not file_exists(file_path):
                messagebox.showerror("错误", f"文件不存在: {file_path}")
                continue

            if is_file_locked(file_path):
                messagebox.showerror("错误", f"文件已被打开: {file_path}")
                continue

            try:
                wb = load_excel(file_path)
                sheets = get_sheet_names(wb)
                del wb

                name = os.path.basename(file_path).rsplit('.', 1)[0]
                item = DataSourceItem.create(name=name, file_path=file_path, sheets=sheets)
                self.store.save_data_source(item)
                success_count += 1
            except Exception as e:
                messagebox.showerror("错误", f"处理失败 {file_path}: {str(e)}")

        if success_count > 0:
            messagebox.showinfo("成功", f"成功导入 {success_count} 个数据源")
            self._refresh_datasource_list(self._find_list_frame("datasource"))

    def _refresh_datasource_list(self, parent_frame):
        """刷新数据源列表"""
        for widget in parent_frame.winfo_children():
            widget.destroy()

        data_sources = self.store.load_data_sources()

        if not data_sources:
            ttk.Label(parent_frame, text="暂无已保存的数据源，请点击上方按钮导入", bootstyle="info").pack()
            return

        # 创建表格（支持多选）
        columns = ("name", "path", "sheets", "status")
        tree = ttk.Treeview(parent_frame, columns=columns, show="headings", height=12, selectmode="extended")

        tree.heading("name", text="名称")
        tree.heading("path", text="文件路径")
        tree.heading("sheets", text="Sheet数")
        tree.heading("status", text="状态")

        tree.column("name", width=150)
        tree.column("path", width=350)
        tree.column("sheets", width=60, anchor=CENTER)
        tree.column("status", width=80, anchor=CENTER)

        for ds in data_sources:
            status = "✅ 正常" if file_exists(ds.file_path) else "⚠️ 不存在"
            tree.insert("", END, values=(ds.name, ds.file_path, len(ds.sheets), status), tags=(ds.id,))

        tree.pack(fill=BOTH, expand=True, pady=5)

        # 按钮
        btn_frame = ttk.Frame(parent_frame)
        btn_frame.pack(fill=X, pady=5)

        def on_delete_selected():
            selected = tree.selection()
            if not selected:
                messagebox.showwarning("提示", "请先选择要删除的数据源")
                return
            ids = [tree.item(item)["tags"][0] for item in selected]
            names = [tree.item(item)["values"][0] for item in selected]
            if messagebox.askyesno("确认", f"确定要删除选中的 {len(ids)} 个数据源吗？\n\n{names}"):
                for ds_id in ids:
                    self.store.delete_data_source(ds_id)
                self._refresh_datasource_list(parent_frame)

        def on_delete_all():
            if not data_sources:
                return
            if messagebox.askyesno("确认", f"确定要删除全部 {len(data_sources)} 个数据源吗？"):
                for ds in data_sources:
                    self.store.delete_data_source(ds.id)
                self._refresh_datasource_list(parent_frame)

        ttk.Button(btn_frame, text="🗑️ 删除选中", bootstyle="danger", command=on_delete_selected).pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="🗑️ 删除全部", bootstyle="danger", command=on_delete_all).pack(side=LEFT, padx=5)

    def _find_list_frame(self, page):
        """找到指定页面的列表frame（递归查找）"""
        target_texts = {
            "datasource": "已保存的数据源",
            "template": "已保存的模板",
            "mapping": "已有映射配置"
        }
        target = target_texts.get(page, "")
        if not target:
            return None

        def find_recursive(parent):
            for widget in parent.winfo_children():
                if isinstance(widget, ttk.LabelFrame):
                    if target in widget.cget("text"):
                        return widget
                # 递归查找子widget
                result = find_recursive(widget)
                if result:
                    return result
            return None

        return find_recursive(self.current_page)

    # ==================== 模板管理页面 ====================

    def _create_template_page(self):
        """创建模板管理页面"""
        frame = ttk.Frame(self.root)

        ttk.Label(frame, text="模板管理", font=("Microsoft YaHei", 16, "bold")).pack(anchor=W, pady=10)
        ttk.Label(frame, text="添加 Excel 模板文件，支持单文件或多文件导入", bootstyle="info").pack(anchor=W, pady=5)

        import_frame = ttk.LabelFrame(frame, text="导入模板")
        import_frame.pack(fill=X, pady=10)

        btn_import_single = ttk.Button(import_frame, text="📂 导入单文件", bootstyle="success",
                                        command=lambda: self._import_template_files(multi=False))
        btn_import_single.pack(side=LEFT, padx=5, pady=10)

        btn_import_multi = ttk.Button(import_frame, text="📂 导入多文件", bootstyle="success",
                                       command=lambda: self._import_template_files(multi=True))
        btn_import_multi.pack(side=LEFT, padx=5, pady=10)

        list_frame = ttk.LabelFrame(frame, text="已保存的模板")
        list_frame.pack(fill=BOTH, expand=True, pady=10)

        self._refresh_template_list(list_frame)

        return frame

    def _import_template_files(self, multi=False):
        """导入模板文件"""
        if multi:
            paths = filedialog.askopenfilenames(title="选择模板文件", filetypes=[("Excel 文件", "*.xlsx *.xlsm")])
        else:
            path = filedialog.askopenfilename(title="选择模板文件", filetypes=[("Excel 文件", "*.xlsx *.xlsm")])
            paths = [path] if path else []

        if not paths:
            return

        success_count = 0
        for file_path in paths:
            if not file_path:
                continue

            if not file_exists(file_path):
                continue
            if is_file_locked(file_path):
                continue

            try:
                wb = load_excel(file_path)
                sheets = get_sheet_names(wb)
                del wb

                name = os.path.basename(file_path).rsplit('.', 1)[0]
                item = TemplateItem.create(name=name, file_path=file_path, sheets=sheets)
                self.store.save_template(item)
                success_count += 1
            except Exception as e:
                messagebox.showerror("错误", f"处理失败: {str(e)}")

        if success_count > 0:
            messagebox.showinfo("成功", f"成功导入 {success_count} 个模板")
            self._refresh_template_list(self._find_list_frame("template"))

    def _refresh_template_list(self, parent_frame):
        for widget in parent_frame.winfo_children():
            widget.destroy()

        templates = self.store.load_templates()

        if not templates:
            ttk.Label(parent_frame, text="暂无已保存的模板，请点击上方按钮导入", bootstyle="info").pack()
            return

        columns = ("name", "path", "sheets", "status")
        tree = ttk.Treeview(parent_frame, columns=columns, show="headings", height=12, selectmode="extended")

        tree.heading("name", text="名称")
        tree.heading("path", text="文件路径")
        tree.heading("sheets", text="Sheet数")
        tree.heading("status", text="状态")

        tree.column("name", width=150)
        tree.column("path", width=350)
        tree.column("sheets", width=60, anchor=CENTER)
        tree.column("status", width=80, anchor=CENTER)

        for t in templates:
            status = "✅ 正常" if file_exists(t.file_path) else "⚠️ 不存在"
            tree.insert("", END, values=(t.name, t.file_path, len(t.sheets), status), tags=(t.id,))

        tree.pack(fill=BOTH, expand=True, pady=5)

        btn_frame = ttk.Frame(parent_frame)
        btn_frame.pack(fill=X, pady=5)

        def on_delete_selected():
            selected = tree.selection()
            if not selected:
                messagebox.showwarning("提示", "请先选择要删除的模板")
                return
            ids = [tree.item(item)["tags"][0] for item in selected]
            names = [tree.item(item)["values"][0] for item in selected]
            if messagebox.askyesno("确认", f"确定要删除选中的 {len(ids)} 个模板吗？\n\n{names}"):
                for t_id in ids:
                    self.store.delete_template(t_id)
                self._refresh_template_list(parent_frame)

        def on_delete_all():
            if not templates:
                return
            if messagebox.askyesno("确认", f"确定要删除全部 {len(templates)} 个模板吗？"):
                for t in templates:
                    self.store.delete_template(t.id)
                self._refresh_template_list(parent_frame)

        ttk.Button(btn_frame, text="🗑️ 删除选中", bootstyle="danger", command=on_delete_selected).pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="🗑️ 删除全部", bootstyle="danger", command=on_delete_all).pack(side=LEFT, padx=5)

    # ==================== 映射配置页面 ====================

    def _create_mapping_page(self):
        """创建映射配置页面"""
        frame = ttk.Frame(self.root)

        ttk.Label(frame, text="映射配置", font=("Microsoft YaHei", 16, "bold")).pack(anchor=W, pady=10)
        ttk.Label(frame, text="导入 JSON 映射配置文件，支持单文件或多文件导入。名称相同则覆盖", bootstyle="info").pack(anchor=W, pady=5)

        # 操作按钮区
        action_frame = ttk.Frame(frame)
        action_frame.pack(fill=X, pady=5)

        # 下载示例
        def download_example():
            example_config = {
                "version": "1.0",
                "default_direction": DEFAULT_DIRECTION,
                "default_write_mode": DEFAULT_WRITE_MODE.value,
                "sheet_mappings": [
                    {
                        "data_source_sheet": "数据源Sheet名称",
                        "template_file": "模板文件名.xlsx",
                        "template_sheet": "模板Sheet名称",
                        "copy_rule": {
                            "source_start": ["A1", "B1", "C1"],
                            "target_start": ["D1", "E1", "F1"],
                            "direction": "horizontal",
                            "length": 12
                        }
                    }
                ]
            }
            path = filedialog.asksaveasfilename(title="保存示例文件", defaultextension=".json",
                                                initialfile=EXAMPLE_MAPPING_FILENAME,
                                                filetypes=[("JSON 文件", "*.json")])
            if path:
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(example_config, f, ensure_ascii=False, indent=2)
                messagebox.showinfo("成功", f"示例文件已保存到:\n{path}")

        ttk.Button(action_frame, text="📥 下载示例文件", command=download_example).pack(side=LEFT, padx=5)

        # 导入按钮
        ttk.Button(action_frame, text="📂 导入单文件", bootstyle="success",
                   command=lambda: self._import_mapping_configs(multi=False)).pack(side=LEFT, padx=5)
        ttk.Button(action_frame, text="📂 导入多文件", bootstyle="success",
                   command=lambda: self._import_mapping_configs(multi=True)).pack(side=LEFT, padx=5)

        # 主内容区：左侧列表 + 右侧详情
        content_frame = ttk.Frame(frame)
        content_frame.pack(fill=BOTH, expand=True, pady=10)

        # 右侧：详情预览（先创建，因为刷新列表时会用到）
        right_frame = ttk.LabelFrame(content_frame, text="映射详情预览")
        right_frame.pack(side=RIGHT, fill=BOTH, expand=True, padx=(5, 0))

        self.detail_text = scrolledtext.ScrolledText(right_frame, font=("Consolas", 9), state="disabled")
        self.detail_text.pack(fill=BOTH, expand=True, padx=5, pady=5)

        # 编辑按钮
        edit_btn_frame = ttk.Frame(right_frame)
        edit_btn_frame.pack(fill=X, padx=5, pady=5)

        self.btn_edit_mapping = ttk.Button(edit_btn_frame, text="✏️ 编辑选中配置",
                                            command=self._edit_selected_mapping, state="disabled")
        self.btn_edit_mapping.pack(side=LEFT, padx=5)

        # 左侧：列表
        left_frame = ttk.LabelFrame(content_frame, text="已有映射配置")
        left_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 5))

        self._refresh_mapping_list(left_frame)

        return frame

    def _import_mapping_configs(self, multi=False):
        """导入映射配置文件"""
        if multi:
            paths = filedialog.askopenfilenames(title="选择映射配置文件", filetypes=[("JSON 文件", "*.json")])
        else:
            path = filedialog.askopenfilename(title="选择映射配置文件", filetypes=[("JSON 文件", "*.json")])
            paths = [path] if path else []

        if not paths:
            return

        success_count = 0
        overwrite_count = 0
        for path in paths:
            if not path:
                continue

            try:
                with open(path, "r", encoding="utf-8") as f:
                    config_data = json.load(f)

                config = MappingConfig.from_dict(config_data)
                name = os.path.basename(path).replace(".json", "")

                # 检查是否已存在同名配置
                existing_configs = self.store.load_mapping_configs()
                existing = next((c for c in existing_configs if c.name == name), None)

                if existing:
                    # 覆盖更新，保留原 id 和创建时间
                    item = MappingConfigItem(
                        id=existing.id,
                        name=name,
                        config=config,
                        created_at=existing.created_at,
                        updated_at=datetime.now().isoformat()
                    )
                    self.store.update_mapping_config(item)
                    overwrite_count += 1
                else:
                    # 新建，使用新 id
                    item = MappingConfigItem.create(name=name, config=config)
                    self.store.save_mapping_config(item)
                    success_count += 1

            except json.JSONDecodeError as e:
                messagebox.showerror("错误", f"JSON 格式错误: {path}\n{str(e)}")
            except Exception as e:
                messagebox.showerror("错误", f"导入失败: {str(e)}")

        msg = f"成功导入 {success_count} 个"
        if overwrite_count > 0:
            msg += f"，覆盖更新 {overwrite_count} 个"
        if success_count > 0 or overwrite_count > 0:
            messagebox.showinfo("成功", msg)
            self._refresh_mapping_list(self._find_list_frame("mapping"))

    def _refresh_mapping_list(self, parent_frame):
        """刷新映射配置列表"""
        for widget in parent_frame.winfo_children():
            widget.destroy()

        configs = self.store.load_mapping_configs()

        if not configs:
            ttk.Label(parent_frame, text="暂无映射配置，请点击上方按钮导入", bootstyle="info").pack()
            return

        columns = ("name", "mappings", "updated")
        tree = ttk.Treeview(parent_frame, columns=columns, show="headings", height=15, selectmode="extended")

        tree.heading("name", text="配置名称")
        tree.heading("mappings", text="映射数")
        tree.heading("updated", text="更新时间")

        tree.column("name", width=180)
        tree.column("mappings", width=60, anchor=CENTER)
        tree.column("updated", width=120)

        for c in configs:
            tree.insert("", END, values=(c.name, len(c.config.sheet_mappings), c.updated_at[:10]), tags=(c.id,))

        tree.pack(fill=BOTH, expand=True, pady=5)

        def on_select(event):
            selected = tree.selection()
            if selected:
                config_id = tree.item(selected[0])["tags"][0]
                config_item = self.store.get_mapping_config(config_id)
                if config_item:
                    self._show_mapping_detail(config_item)
                    self.btn_edit_mapping.config(state="normal")
            else:
                self.btn_edit_mapping.config(state="disabled")

        tree.bind("<<TreeviewSelect>>", on_select)

        # 默认选中第一个
        if configs:
            tree.selection_set(tree.get_children()[0])
            config_item = self.store.get_mapping_config(tree.item(tree.get_children()[0])["tags"][0])
            if config_item:
                self._show_mapping_detail(config_item)
                self.btn_edit_mapping.config(state="normal")

        # 按钮
        btn_frame = ttk.Frame(parent_frame)
        btn_frame.pack(fill=X, pady=5)

        def on_delete_selected():
            selected = tree.selection()
            if not selected:
                messagebox.showwarning("提示", "请先选择要删除的映射配置")
                return
            ids = [tree.item(item)["tags"][0] for item in selected]
            names = [tree.item(item)["values"][0] for item in selected]
            if messagebox.askyesno("确认", f"确定要删除选中的 {len(ids)} 个映射配置吗？\n\n{names}"):
                for config_id in ids:
                    self.store.delete_mapping_config(config_id)
                # 清空详情
                self.detail_text.config(state="normal")
                self.detail_text.delete("1.0", END)
                self.detail_text.insert(END, "已删除，请选择其他配置")
                self.detail_text.config(state="disabled")
                self.btn_edit_mapping.config(state="disabled")
                self._refresh_mapping_list(parent_frame)

        def on_delete_all():
            if not configs:
                return
            if messagebox.askyesno("确认", f"确定要删除全部 {len(configs)} 个映射配置吗？"):
                for c in configs:
                    self.store.delete_mapping_config(c.id)
                self.detail_text.config(state="normal")
                self.detail_text.delete("1.0", END)
                self.detail_text.insert(END, "已全部删除")
                self.detail_text.config(state="disabled")
                self.btn_edit_mapping.config(state="disabled")
                self._refresh_mapping_list(parent_frame)

        ttk.Button(btn_frame, text="🗑️ 删除选中", bootstyle="danger", command=on_delete_selected).pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="🗑️ 删除全部", bootstyle="danger", command=on_delete_all).pack(side=LEFT, padx=5)

    def _show_mapping_detail(self, config_item):
        """显示映射配置详情"""
        self.detail_text.config(state="normal")
        self.detail_text.delete("1.0", END)

        self.detail_text.insert(END, f"配置名称: {config_item.name}\n")
        self.detail_text.insert(END, f"更新时间: {config_item.updated_at}\n")
        self.detail_text.insert(END, f"映射数量: {len(config_item.config.sheet_mappings)} 个\n")
        self.detail_text.insert(END, "\n" + "="*50 + "\n\n")

        # 获取数据源和模板信息用于显示
        data_sources = self.store.load_data_sources()
        templates = self.store.load_templates()

        for i, mapping in enumerate(config_item.config.sheet_mappings):
            rule = mapping.copy_rule
            # 查找对应的数据源名称
            ds_name = "未知数据源"
            for ds in data_sources:
                if mapping.data_source_sheet in ds.sheets:
                    ds_name = ds.name
                    break

            # 查找模板sheet
            tmpl_sheet = mapping.template_sheet if mapping.template_sheet else "默认"

            self.detail_text.insert(END, f"【映射 {i+1}】\n")
            self.detail_text.insert(END, f"  数据源: {ds_name} > {mapping.data_source_sheet}\n")
            self.detail_text.insert(END, f"  → 模板: {mapping.template_file} > {tmpl_sheet}\n")
            self.detail_text.insert(END, f"  规则: {rule.source_start} → {rule.target_start}, {rule.direction.value}, {rule.length}格\n")
            self.detail_text.insert(END, f"  写入: {config_item.config.write_mode.value}\n\n")

        self.detail_text.insert(END, "="*50 + "\n")
        self.detail_text.insert(END, "【完整 JSON 配置】\n")
        self.detail_text.insert(END, json.dumps(config_item.config.to_dict(), ensure_ascii=False, indent=2))

        self.detail_text.config(state="disabled")

    def _edit_selected_mapping(self):
        """编辑选中的映射配置"""
        list_frame = self._find_list_frame("mapping")
        if not list_frame:
            messagebox.showwarning("提示", "请先选择一个映射配置")
            return

        # 找到treeview和选中的项
        tree = None
        for widget in list_frame.winfo_children():
            if isinstance(widget, ttk.Treeview):
                tree = widget
                break

        if not tree or not tree.selection():
            messagebox.showwarning("提示", "请先在列表中选择一个映射配置")
            return

        config_id = tree.item(tree.selection()[0])["tags"][0]
        config_item = self.store.get_mapping_config(config_id)
        if not config_item:
            return

        self._open_mapping_editor(config_item)

    def _open_mapping_editor(self, config_item):
        """打开映射配置编辑器"""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"编辑映射配置")
        dialog.geometry("1100x750")
        dialog.transient(self.root)

        # 存储当前是否放大状态
        is_expanded = {"value": False}

        main_frame = ttk.Frame(dialog)
        main_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)

        # 顶部按钮栏
        top_btn_frame = ttk.Frame(main_frame)
        top_btn_frame.pack(fill=X, pady=(0, 5))

        # 放大/缩小按钮
        def toggle_expand():
            if is_expanded["value"]:
                dialog.state('normal')
                dialog.geometry("1100x750")
                toggle_btn.config(text="🔍 放大")
                is_expanded["value"] = False
            else:
                # 最大化窗口
                dialog.state('zoomed')
                toggle_btn.config(text="🔍 缩小")
                is_expanded["value"] = True

        toggle_btn = ttk.Button(top_btn_frame, text="🔍 放大", command=toggle_expand, bootstyle="info")
        toggle_btn.pack(side=LEFT, padx=5)

        # 基本信息（名称只读）
        info_frame = ttk.LabelFrame(main_frame, text="基本信息")
        info_frame.pack(fill=X, pady=5)

        ttk.Label(info_frame, text="配置名称:").grid(row=0, column=0, sticky=W, padx=5, pady=5)
        ttk.Label(info_frame, text=config_item.name, bootstyle="info").grid(row=0, column=1, sticky=W, padx=5, pady=5)

        # 写入模式选择
        ttk.Label(info_frame, text="写入模式:").grid(row=0, column=2, sticky=W, padx=5, pady=5)
        write_mode_var = tk.StringVar(value=config_item.config.write_mode.value)
        write_mode_combo = ttk.Combobox(
            info_frame,
            values=["skip_nonempty", "overwrite"],
            textvariable=write_mode_var,
            width=15,
            state="readonly"
        )
        write_mode_combo.grid(row=0, column=3, sticky=W, padx=5, pady=5)
        ttk.Label(info_frame, text="(skip:只写空行, overwrite:直接覆盖)", bootstyle="secondary").grid(row=0, column=4, sticky=W, padx=5, pady=5)

        # 映射列表
        mappings_frame = ttk.LabelFrame(main_frame, text="映射列表")
        mappings_frame.pack(fill=BOTH, expand=True, pady=5)

        # 创建带滚动条的canvas
        canvas = tk.Canvas(mappings_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(mappings_frame, orient="vertical", command=canvas.yview)
        scrollable = ttk.Frame(canvas)

        scrollable.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # 鼠标滚轮支持
        def on_mousewheel(event):
            try:
                if canvas.winfo_exists():
                    canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except:
                pass
        canvas.bind_all("<MouseWheel>", on_mousewheel)

        mapping_vars = []
        mapping_frames = []

        # 数据源和模板列表
        data_sources = self.store.load_data_sources()
        templates = self.store.load_templates()

        # 找到当前映射对应的数据源
        def find_ds_by_sheet(sheet_name):
            for ds in data_sources:
                if sheet_name in ds.sheets:
                    return ds
            return None

        # 找到当前映射对应的模板
        def find_tmpl_by_file(file_name):
            for t in templates:
                if os.path.basename(t.file_path) == file_name:
                    return t
            return None

        # 创建单个映射控件
        def create_mapping_frame(parent, index, mapping=None):
            mf = ttk.LabelFrame(parent, text=f"映射 {index + 1}")
            mf.pack(fill=X, pady=5, padx=5)

            mv = {"index": index, "frame": mf}

            # 如果没有传入mapping，使用默认值
            if mapping is None:
                mapping = type('Mapping', (), {
                    'data_source_sheet': '',
                    'template_file': templates[0].name if templates else '',
                    'template_sheet': '',
                    'copy_rule': type('Rule', (), {
                        'source_start': 'A1',
                        'target_start': 'A1',
                        'direction': 'horizontal',
                        'length': 10
                    })()
                })()

            current_ds = find_ds_by_sheet(mapping.data_source_sheet)
            current_tmpl = find_tmpl_by_file(mapping.template_file)

            # 删除按钮
            def delete_this():
                mf.destroy()
                mapping_vars[:] = [v for v in mapping_vars if v.get("frame") != mf]
                # 重新编号
                for i, v in enumerate(mapping_vars):
                    v["frame"].configure(text=f"映射 {i + 1}")
                mapping_vars[:] = [dict(list(v.items()) + [("index", i)]) if "index" in v else v for i, v in enumerate(mapping_vars)]

            del_btn = ttk.Button(mf, text="🗑️ 删除", command=delete_this, bootstyle="danger")
            del_btn.pack(side=RIGHT, anchor=NE, padx=5, pady=2)

            # 行0: 数据源选择
            row0_frame = ttk.Frame(mf)
            row0_frame.pack(fill=X, pady=2)

            ttk.Label(row0_frame, text="数据源:").pack(side=LEFT, padx=(0, 2))
            mv["ds"] = tk.StringVar(value=current_ds.name if current_ds else "")
            ds_combo = ttk.Combobox(row0_frame, values=[ds.name for ds in data_sources],
                                    textvariable=mv["ds"], width=20, state="readonly")
            ds_combo.pack(side=LEFT, padx=2)

            ttk.Label(row0_frame, text="Sheet:").pack(side=LEFT, padx=(10, 2))
            mv["ds_sheet"] = tk.StringVar(value=mapping.data_source_sheet)
            ds_sheet_combo = ttk.Combobox(row0_frame, values=current_ds.sheets if current_ds else [],
                                           textvariable=mv["ds_sheet"], width=15, state="readonly")
            ds_sheet_combo.pack(side=LEFT, padx=2)

            def update_ds_sheets_combo(idx, combo, *args):
                ds_name = mapping_vars[idx]["ds"].get()
                ds = next((d for d in data_sources if d.name == ds_name), None)
                if ds:
                    combo["values"] = ds.sheets
                    if mapping_vars[idx]["ds_sheet"].get() not in ds.sheets:
                        mapping_vars[idx]["ds_sheet"].set(ds.sheets[0] if ds.sheets else "")

            ds_combo.bind("<<ComboboxSelected>>", lambda e, idx=index, combo=ds_sheet_combo: update_ds_sheets_combo(idx, combo))

            # 行1: 模板选择
            row1_frame = ttk.Frame(mf)
            row1_frame.pack(fill=X, pady=2)

            ttk.Label(row1_frame, text="模板文件:").pack(side=LEFT, padx=(0, 2))
            mv["template"] = tk.StringVar(value=mapping.template_file)
            t_combo = ttk.Combobox(row1_frame, values=[os.path.basename(t.file_path) for t in templates],
                                   textvariable=mv["template"], width=25, state="readonly")
            t_combo.pack(side=LEFT, padx=2)

            ttk.Label(row1_frame, text="模板Sheet:").pack(side=LEFT, padx=(10, 2))
            mv["template_sheet"] = tk.StringVar(value=mapping.template_sheet)
            t_sheet_combo = ttk.Combobox(row1_frame, values=current_tmpl.sheets if current_tmpl else [],
                                         textvariable=mv["template_sheet"], width=15, state="readonly")
            t_sheet_combo.pack(side=LEFT, padx=2)

            def update_template_sheets_combo(idx, combo, *args):
                t_name = mapping_vars[idx]["template"].get()
                t = find_tmpl_by_file(t_name)
                if t:
                    combo["values"] = t.sheets
                    if mapping_vars[idx]["template_sheet"].get() not in t.sheets:
                        mapping_vars[idx]["template_sheet"].set(t.sheets[0] if t.sheets else "")

            t_combo.bind("<<ComboboxSelected>>", lambda e, idx=index, combo=t_sheet_combo: update_template_sheets_combo(idx, combo))

            # 行2: 源起始和目标起始（使用大文本框）
            row2_frame = ttk.Frame(mf)
            row2_frame.pack(fill=X, pady=2)

            ttk.Label(row2_frame, text="源起始:").pack(side=LEFT, padx=(0, 2))
            src_val = ", ".join(mapping.copy_rule.source_start) if isinstance(mapping.copy_rule.source_start, list) else mapping.copy_rule.source_start
            src_text = scrolledtext.ScrolledText(row2_frame, width=35, height=3, wrap=tk.WORD)
            src_text.insert("1.0", src_val)
            src_text.pack(side=LEFT, padx=2, fill=X, expand=True)
            mv["source_start"] = src_text

            ttk.Label(row2_frame, text="目标起始:").pack(side=LEFT, padx=(10, 2))
            tgt_val = ", ".join(mapping.copy_rule.target_start) if isinstance(mapping.copy_rule.target_start, list) else mapping.copy_rule.target_start
            tgt_text = scrolledtext.ScrolledText(row2_frame, width=35, height=3, wrap=tk.WORD)
            tgt_text.insert("1.0", tgt_val)
            tgt_text.pack(side=LEFT, padx=2, fill=X, expand=True)
            mv["target_start"] = tgt_text

            # 行3: 方向和长度
            row3_frame = ttk.Frame(mf)
            row3_frame.pack(fill=X, pady=2)

            ttk.Label(row3_frame, text="方向:").pack(side=LEFT, padx=(0, 2))
            mv["direction"] = tk.StringVar(value=mapping.copy_rule.direction.value)
            ttk.Combobox(row3_frame, values=["horizontal", "vertical"], textvariable=mv["direction"], width=12, state="readonly").pack(side=LEFT, padx=2)

            ttk.Label(row3_frame, text="长度:").pack(side=LEFT, padx=(10, 2))
            mv["length"] = tk.IntVar(value=mapping.copy_rule.length)
            ttk.Entry(row3_frame, textvariable=mv["length"], width=10).pack(side=LEFT, padx=2)

            return mv

        # 创建已有的映射
        for i, mapping in enumerate(config_item.config.sheet_mappings):
            mv = create_mapping_frame(scrollable, i, mapping)
            mapping_vars.append(mv)

        # 新增映射按钮
        def add_new_mapping():
            idx = len(mapping_vars)
            mv = create_mapping_frame(scrollable, idx)
            mapping_vars.append(mv)
            # 滚动到底部
            dialog.update_idletasks()
            canvas.yview_moveto(1.0)

        add_btn_frame = ttk.Frame(mappings_frame)
        add_btn_frame.pack(fill=X, pady=5)

        ttk.Button(add_btn_frame, text="➕ 添加映射", command=add_new_mapping, bootstyle="success").pack(side=LEFT, padx=5)

        canvas.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        # 底部按钮
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill=X, pady=10)

        def do_save():
            sheet_mappings = []
            for mv in mapping_vars:
                if "ds_sheet" not in mv:  # 跳过已删除的
                    continue
                # 规范化输入：处理全角逗号、中文逗号等
                # 从大文本框获取内容
                src = normalize_list_string(mv["source_start"].get("1.0", tk.END).strip())
                tgt = normalize_list_string(mv["target_start"].get("1.0", tk.END).strip())
                direction = normalize_string(mv["direction"].get())
                # 长度：取第一个有效数字
                length_str = normalize_string(str(mv["length"].get()))
                try:
                    length = int(length_str.split(",")[0].strip())
                except (ValueError, TypeError):
                    length = 10

                sheet_mappings.append({
                    "data_source_sheet": mv["ds_sheet"].get(),
                    "template_file": mv["template"].get(),
                    "template_sheet": mv["template_sheet"].get(),
                    "copy_rule": {
                        "source_start": [x.strip() for x in src.split(",")] if "," in src else src,
                        "target_start": [x.strip() for x in tgt.split(",")] if "," in tgt else tgt,
                        "direction": direction,
                        "length": length
                    }
                })

            # 规范化写入模式
            write_mode = normalize_string(write_mode_var.get())

            new_config = MappingConfig.from_dict({
                "version": "1.0",
                "default_direction": config_item.config.default_direction.value,
                "write_mode": write_mode,
                "sheet_mappings": sheet_mappings
            })

            updated = MappingConfigItem(
                id=config_item.id,
                name=config_item.name,
                config=new_config,
                created_at=config_item.created_at,
                updated_at=datetime.now().isoformat()
            )

            self.store.update_mapping_config(updated)
            messagebox.showinfo("成功", "保存成功")
            dialog.destroy()
            self._refresh_mapping_list(self._find_list_frame("mapping"))

        ttk.Button(btn_frame, text="💾 保存", bootstyle="success", command=do_save).pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=LEFT)

        # 窗口关闭时解绑鼠标滚轮事件
        def on_dialog_close():
            try:
                canvas.unbind_all("<MouseWheel>")
            except:
                pass
            dialog.destroy()
        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)

    def _delete_selected_mapping(self):
        """删除选中的映射配置"""
        list_frame = self._find_list_frame("mapping")
        if not list_frame:
            return

        tree = None
        for widget in list_frame.winfo_children():
            if isinstance(widget, ttk.Treeview):
                tree = widget
                break

        if not tree or not tree.selection():
            messagebox.showwarning("提示", "请先选择一个映射配置")
            return

        config_id = tree.item(tree.selection()[0])["tags"][0]
        if messagebox.askyesno("确认", "确定要删除这个映射配置吗？"):
            self.store.delete_mapping_config(config_id)
            self.detail_text.config(state="normal")
            self.detail_text.delete("1.0", END)
            self.detail_text.insert(END, "已删除，请选择其他配置")
            self.detail_text.config(state="disabled")
            self.btn_edit_mapping.config(state="disabled")
            self._refresh_mapping_list(list_frame)

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

        # 只选择映射配置
        select_frame = ttk.LabelFrame(frame, text="选择映射配置（数据源和模板将根据配置自动匹配）")
        select_frame.pack(fill=X, pady=10)

        cfg_frame = ttk.Frame(select_frame)
        cfg_frame.pack(fill=X, pady=5)
        ttk.Label(cfg_frame, text="映射配置:").pack(side=LEFT, padx=5)
        cfg_var = tk.StringVar(value=configs[0].name)
        cfg_combo = ttk.Combobox(cfg_frame, values=[c.name for c in configs],
                                  textvariable=cfg_var, state="readonly", width=40)
        cfg_combo.pack(side=LEFT, padx=5)

        # 预览区
        preview_frame = ttk.LabelFrame(frame, text="迁移预览")
        preview_frame.pack(fill=BOTH, expand=True, pady=10)

        preview_text = scrolledtext.ScrolledText(preview_frame, height=12, font=("Consolas", 9), state="disabled")
        preview_text.pack(fill=BOTH, expand=True, padx=5, pady=5)

        # 结果区
        result_frame = ttk.LabelFrame(frame, text="执行结果")
        result_frame.pack(fill=X, pady=10)

        result_text = scrolledtext.ScrolledText(result_frame, height=8, font=("Consolas", 9), state="disabled")
        result_text.pack(fill=X, padx=5, pady=5)

        def update_preview(*args):
            preview_text.config(state="normal")
            preview_text.delete("1.0", END)

            cfg_name = cfg_var.get()
            cfg = next((c for c in configs if c.name == cfg_name), None)

            if not cfg:
                preview_text.insert("1.0", "请选择映射配置")
                preview_text.config(state="disabled")
                return

            # 构建数据源sheet到数据源的映射
            ds_sheet_to_ds = {}
            for ds in data_sources:
                for sheet in ds.sheets:
                    ds_sheet_to_ds[sheet] = ds

            template_name_map = {os.path.basename(t.file_path): t for t in templates}

            matched, unmatched_src, unmatched_tgt, used_ds = [], [], [], set()

            for m in cfg.config.sheet_mappings:
                ds = ds_sheet_to_ds.get(m.data_source_sheet)
                if not ds:
                    unmatched_src.append(m.data_source_sheet)
                    continue
                if m.template_file not in template_name_map:
                    unmatched_tgt.append(m.template_file)
                    continue
                matched.append((ds, m))
                used_ds.add(ds.name)

            if matched:
                preview_text.insert(END, f"✅ 将执行 {len(matched)} 个迁移:\n\n")
                for ds, m in matched:
                    r = m.copy_rule
                    tmpl_sheet = m.template_sheet if m.template_sheet else "默认"
                    preview_text.insert(END, f"• {ds.name} > {m.data_source_sheet}\n")
                    preview_text.insert(END, f"  → {m.template_file} > {tmpl_sheet}\n")
                    preview_text.insert(END, f"  规则: {r.source_start} → {r.target_start}, {r.direction.value}, {r.length}格\n\n")

                # 显示将使用的数据源和模板
                preview_text.insert(END, "="*50 + "\n")
                preview_text.insert(END, f"将使用 {len(used_ds)} 个数据源: {', '.join(used_ds)}\n")
                t_used = set(os.path.basename(t.file_path) for _, m in matched for t in templates if os.path.basename(t.file_path) == m.template_file)
                preview_text.insert(END, f"将使用 {len(t_used)} 个模板: {', '.join(t_used)}\n")

            if unmatched_src:
                preview_text.insert(END, f"\n⚠️ 数据源中不存在 ({len(unmatched_src)}): {', '.join(unmatched_src)}\n")
            if unmatched_tgt:
                preview_text.insert(END, f"⚠️ 未找到模板 ({len(unmatched_tgt)}): {', '.join(unmatched_tgt)}\n")

            preview_text.config(state="disabled")

        cfg_combo.bind("<<ComboboxSelected>>", update_preview)

        # 执行按钮
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=X, pady=5)

        def execute():
            result_text.config(state="normal")
            result_text.delete("1.0", END)
            result_text.update()

            cfg = next((c for c in configs if c.name == cfg_var.get()), None)
            if not cfg:
                result_text.insert("1.0", "请选择映射配置")
                result_text.config(state="disabled")
                return

            # 自动匹配数据源和模板
            ds_sheet_to_ds = {}
            for ds in data_sources:
                for sheet in ds.sheets:
                    ds_sheet_to_ds[sheet] = ds

            template_name_map = {os.path.basename(t.file_path): t for t in templates}

            # 按模板文件分组映射
            template_groups = {}
            for m in cfg.config.sheet_mappings:
                ds = ds_sheet_to_ds.get(m.data_source_sheet)
                t_item = template_name_map.get(m.template_file)
                if not ds or not t_item:
                    continue
                if m.template_file not in template_groups:
                    template_groups[m.template_file] = {"ds": ds, "t_item": t_item, "mappings": []}
                template_groups[m.template_file]["mappings"].append(m)

            if not template_groups:
                result_text.insert("1.0", "❌ 无法匹配任何有效的数据源和模板组合\n请确保数据源和模板已正确导入")
                result_text.config(state="disabled")
                return

            try:
                all_results = []

                for t_file, group in template_groups.items():
                    ds = group["ds"]
                    t_item = group["t_item"]
                    mappings = group["mappings"]

                    if not file_exists(ds.file_path):
                        result_text.insert(END, f"❌ 数据源文件不存在: {ds.file_path}\n")
                        continue
                    if not file_exists(t_item.file_path):
                        result_text.insert(END, f"❌ 模板文件不存在: {t_item.file_path}\n")
                        continue

                    source_wb = load_excel(ds.file_path, data_only=True)
                    target_path = t_item.file_path

                    if is_file_locked(target_path):
                        target_path = _generate_copy_path(target_path)
                        shutil.copy2(t_item.file_path, target_path)
                        result_text.insert(END, f"📁 原始文件被占用，已生成副本: {target_path}\n")
                        result_text.update()

                    target_wb = load_excel(target_path)
                    copier = DataCopier(write_mode=WriteMode.OVERWRITE)

                    for mapping in mappings:
                        result = copier.copy(source_wb, target_wb, mapping)
                        all_results.append(result)

                    target_wb.save(target_path)

                success = [r for r in all_results if r.success]
                failed = [r for r in all_results if not r.success]

                result_text.insert(END, "="*40 + "\n")
                if success:
                    result_text.insert(END, f"✅ 成功: {len(success)} 个\n")
                    for r in success:
                        # 显示详细信息
                        ds_name = ""
                        for ds in data_sources:
                            if r.source_sheet in ds.sheets:
                                ds_name = ds.name
                                break
                        result_text.insert(END, f"  - {ds_name} > {r.source_sheet}\n")
                        result_text.insert(END, f"    → {r.message}\n")
                if failed:
                    result_text.insert(END, f"\n⚠️ 跳过: {len(failed)} 个\n")
                    for r in failed:
                        result_text.insert(END, f"  - {r.source_sheet}: {r.message}\n")
                result_text.insert(END, f"\n{'='*40}\n总计: {len(all_results)} | 成功: {len(success)} | 跳过: {len(failed)}\n")

            except PermissionError:
                result_text.insert("1.0", "⚠️ 文件已打开，请关闭后重试")
            except Exception as e:
                result_text.insert("1.0", f"❌ 迁移失败: {str(e)}")

            result_text.config(state="disabled")

        ttk.Button(btn_frame, text="🚀 执行迁移", bootstyle="success", command=execute).pack(side=LEFT, padx=5)

        update_preview()

        return frame


def _generate_copy_path(original_path: str) -> str:
    """生成副本文件路径"""
    directory = os.path.dirname(original_path)
    filename = os.path.basename(original_path)
    name, ext = os.path.splitext(filename)
    return os.path.join(directory, f"{name}_副本{ext}")


def main():
    ExcelMigratorApp()


if __name__ == "__main__":
    main()
