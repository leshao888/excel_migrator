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
from core.enums import Direction, WriteMode
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
        self.root.title("Excel 数据迁移工具 v1.0")
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
        version_label = ttk.Label(sidebar, text="v1.0", bootstyle="secondary")
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
        messagebox.showinfo("关于", "Excel 数据迁移工具 v1.0\n\n用于 Excel 数据迁移的工具软件")

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

        # 创建表格
        columns = ("name", "path", "sheets", "status")
        tree = ttk.Treeview(parent_frame, columns=columns, show="headings", height=12)

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

        def on_delete():
            selected = tree.selection()
            if selected:
                ds_id = tree.item(selected[0])["tags"][0]
                if messagebox.askyesno("确认", "确定要删除这个数据源吗？"):
                    self.store.delete_data_source(ds_id)
                    self._refresh_datasource_list(parent_frame)

        ttk.Button(btn_frame, text="🗑️ 删除", bootstyle="danger", command=on_delete).pack(side=LEFT, padx=5)

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
        tree = ttk.Treeview(parent_frame, columns=columns, show="headings", height=12)

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

        def on_delete():
            selected = tree.selection()
            if selected:
                t_id = tree.item(selected[0])["tags"][0]
                if messagebox.askyesno("确认", "确定要删除这个模板吗？"):
                    self.store.delete_template(t_id)
                    self._refresh_template_list(parent_frame)

        ttk.Button(btn_frame, text="🗑️ 删除", bootstyle="danger", command=on_delete).pack(side=LEFT, padx=5)

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

        ttk.Button(edit_btn_frame, text="🗑️ 删除选中配置", bootstyle="danger",
                   command=self._delete_selected_mapping).pack(side=LEFT, padx=5)

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
        tree = ttk.Treeview(parent_frame, columns=columns, show="headings", height=15)

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

        tree.bind("<<TreeviewSelect>>", on_select)

        # 默认选中第一个
        if configs:
            tree.selection_set(tree.get_children()[0])
            config_item = self.store.get_mapping_config(tree.item(tree.get_children()[0])["tags"][0])
            if config_item:
                self._show_mapping_detail(config_item)
                self.btn_edit_mapping.config(state="normal")

    def _show_mapping_detail(self, config_item):
        """显示映射配置详情"""
        self.detail_text.config(state="normal")
        self.detail_text.delete("1.0", END)

        self.detail_text.insert(END, f"配置名称: {config_item.name}\n")
        self.detail_text.insert(END, f"更新时间: {config_item.updated_at}\n")
        self.detail_text.insert(END, f"映射数量: {len(config_item.config.sheet_mappings)} 个\n")
        self.detail_text.insert(END, "\n" + "="*50 + "\n\n")

        for i, mapping in enumerate(config_item.config.sheet_mappings):
            rule = mapping.copy_rule
            self.detail_text.insert(END, f"【映射 {i+1}】\n")
            self.detail_text.insert(END, f"  数据源 Sheet: {mapping.data_source_sheet}\n")
            self.detail_text.insert(END, f"  模板文件: {mapping.template_file}\n")
            self.detail_text.insert(END, f"  源起始: {rule.source_start}\n")
            self.detail_text.insert(END, f"  目标起始: {rule.target_start}\n")
            self.detail_text.insert(END, f"  方向: {rule.direction.value}\n")
            self.detail_text.insert(END, f"  长度: {rule.length}\n\n")

        self.detail_text.insert(END, "="*50 + "\n")
        self.detail_text.insert(END, "【完整 JSON 配置】\n")
        self.detail_text.insert(END, json.dumps(config_item.config.to_dict(), ensure_ascii=False, indent=2))

        self.detail_text.config(state="disabled")

    def _edit_selected_mapping(self):
        """编辑选中的映射配置"""
        list_frame = self._find_list_frame("mapping")
        if not list_frame:
            return

        # 找到treeview
        tree = None
        for widget in list_frame.winfo_children():
            if isinstance(widget, ttk.Treeview):
                tree = widget
                break

        if not tree or not tree.selection():
            messagebox.showwarning("提示", "请先选择一个映射配置")
            return

        config_id = tree.item(tree.selection()[0])["tags"][0]
        config_item = self.store.get_mapping_config(config_id)
        if not config_item:
            return

        self._open_mapping_editor(config_item)

    def _open_mapping_editor(self, config_item):
        """打开映射配置编辑器"""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"编辑映射配置 - {config_item.name}")
        dialog.geometry("900x600")
        dialog.transient(self.root)

        main_frame = ttk.Frame(dialog)
        main_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)

        # 基本信息（名称可编辑）
        info_frame = ttk.LabelFrame(main_frame, text="基本信息")
        info_frame.pack(fill=X, pady=5)

        ttk.Label(info_frame, text="配置名称:").grid(row=0, column=0, sticky=W, padx=5, pady=5)
        name_var = tk.StringVar(value=config_item.name)
        ttk.Entry(info_frame, textvariable=name_var, width=40).grid(row=0, column=1, sticky=W, padx=5, pady=5)

        ttk.Label(info_frame, text="JSON路径:").grid(row=0, column=2, sticky=W, padx=5, pady=5)
        path_label = ttk.Label(info_frame, text="存储在配置文件中，不可编辑", bootstyle="info")
        path_label.grid(row=0, column=3, sticky=W, padx=5, pady=5)

        # 映射列表
        mappings_frame = ttk.LabelFrame(main_frame, text="映射列表")
        mappings_frame.pack(fill=BOTH, expand=True, pady=5)

        # 创建带滚动条的canvas
        canvas = tk.Canvas(mappings_frame)
        scrollbar = ttk.Scrollbar(mappings_frame, orient="vertical", command=canvas.yview)
        scrollable = ttk.Frame(canvas)

        scrollable.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        mapping_vars = []

        # 数据源和模板列表
        data_sources = self.store.load_data_sources()
        templates = self.store.load_templates()

        for i, mapping in enumerate(config_item.config.sheet_mappings):
            mf = ttk.LabelFrame(scrollable, text=f"映射 {i+1}", padding=5)
            mf.pack(fill=X, pady=3, padx=5)

            mv = {"index": i}

            # 数据源Sheet
            ttk.Label(mf, text="数据源Sheet:").grid(row=0, column=0, sticky=W, padx=2, pady=2)
            ds_sheets = [s for ds in data_sources for s in ds.sheets] if data_sources else []
            mv["ds_sheet"] = tk.StringVar(value=mapping.data_source_sheet)
            ttk.Combobox(mf, values=ds_sheets, textvariable=mv["ds_sheet"], width=20, state="readonly").grid(row=0, column=1, sticky=W, padx=2, pady=2)

            # 模板文件
            ttk.Label(mf, text="模板文件:").grid(row=0, column=2, sticky=W, padx=2, pady=2)
            t_files = [os.path.basename(t.file_path) for t in templates] if templates else []
            mv["template"] = tk.StringVar(value=mapping.template_file)
            ttk.Combobox(mf, values=t_files, textvariable=mv["template"], width=20, state="readonly").grid(row=0, column=3, sticky=W, padx=2, pady=2)

            # 源起始
            ttk.Label(mf, text="源起始:").grid(row=1, column=0, sticky=W, padx=2, pady=2)
            src_val = ", ".join(mapping.copy_rule.source_start) if isinstance(mapping.copy_rule.source_start, list) else mapping.copy_rule.source_start
            mv["source_start"] = tk.StringVar(value=src_val)
            ttk.Entry(mf, textvariable=mv["source_start"], width=25).grid(row=1, column=1, sticky=W, padx=2, pady=2)

            # 目标起始
            ttk.Label(mf, text="目标起始:").grid(row=1, column=2, sticky=W, padx=2, pady=2)
            tgt_val = ", ".join(mapping.copy_rule.target_start) if isinstance(mapping.copy_rule.target_start, list) else mapping.copy_rule.target_start
            mv["target_start"] = tk.StringVar(value=tgt_val)
            ttk.Entry(mf, textvariable=mv["target_start"], width=25).grid(row=1, column=3, sticky=W, padx=2, pady=2)

            # 方向和长度
            ttk.Label(mf, text="方向:").grid(row=2, column=0, sticky=W, padx=2, pady=2)
            mv["direction"] = tk.StringVar(value=mapping.copy_rule.direction.value)
            ttk.Combobox(mf, values=["horizontal", "vertical"], textvariable=mv["direction"], width=18, state="readonly").grid(row=2, column=1, sticky=W, padx=2, pady=2)

            ttk.Label(mf, text="长度:").grid(row=2, column=2, sticky=W, padx=2, pady=2)
            mv["length"] = tk.IntVar(value=mapping.copy_rule.length)
            ttk.Entry(mf, textvariable=mv["length"], width=25).grid(row=2, column=3, sticky=W, padx=2, pady=2)

            mapping_vars.append(mv)

        canvas.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        # 按钮
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill=X, pady=10)

        def do_save():
            sheet_mappings = []
            for mv in mapping_vars:
                src = mv["source_start"].get().strip()
                tgt = mv["target_start"].get().strip()
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

            updated = MappingConfigItem(
                id=config_item.id,
                name=name_var.get().strip(),
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

        # 选择区
        select_frame = ttk.LabelFrame(frame, text="选择数据源、模板和映射配置")
        select_frame.pack(fill=X, pady=10)

        # 数据源
        ds_frame = ttk.Frame(select_frame)
        ds_frame.pack(fill=X, pady=5)
        ttk.Label(ds_frame, text="数据源:").pack(side=LEFT, padx=5)
        ds_var = tk.StringVar(value=data_sources[0].name)
        ds_combo = ttk.Combobox(ds_frame, values=[ds.name for ds in data_sources],
                                textvariable=ds_var, state="readonly", width=30)
        ds_combo.pack(side=LEFT, padx=5)

        # 模板
        t_frame = ttk.Frame(select_frame)
        t_frame.pack(fill=X, pady=5)
        ttk.Label(t_frame, text="模板:").pack(side=LEFT, padx=5)
        t_var = tk.StringVar(value=templates[0].name)
        t_combo = ttk.Combobox(t_frame, values=[t.name for t in templates],
                                textvariable=t_var, state="readonly", width=30)
        t_combo.pack(side=LEFT, padx=5)

        # 映射配置
        cfg_frame = ttk.Frame(select_frame)
        cfg_frame.pack(fill=X, pady=5)
        ttk.Label(cfg_frame, text="映射配置:").pack(side=LEFT, padx=5)
        cfg_var = tk.StringVar(value=configs[0].name)
        cfg_combo = ttk.Combobox(cfg_frame, values=[c.name for c in configs],
                                  textvariable=cfg_var, state="readonly", width=30)
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

            ds_name = ds_var.get()
            cfg_name = cfg_var.get()

            ds = next((d for d in data_sources if d.name == ds_name), None)
            cfg = next((c for c in configs if c.name == cfg_name), None)

            if not ds or not cfg:
                preview_text.insert("1.0", "请选择数据源和映射配置")
                preview_text.config(state="disabled")
                return

            template_name_map = {os.path.basename(t.file_path): t for t in templates}

            matched, unmatched_src, unmatched_tgt = [], [], []

            for m in cfg.config.sheet_mappings:
                if m.data_source_sheet not in ds.sheets:
                    unmatched_src.append(m.data_source_sheet)
                    continue
                if m.template_file not in template_name_map:
                    unmatched_tgt.append(m.template_file)
                    continue
                matched.append(m)

            if matched:
                preview_text.insert(END, f"✅ 将迁移 {len(matched)} 个 Sheet:\n\n")
                for m in matched:
                    r = m.copy_rule
                    preview_text.insert(END, f"• {m.data_source_sheet} → {m.template_file}\n")
                    preview_text.insert(END, f"  {r.source_start} → {r.target_start}, {r.direction.value}, {r.length}格\n\n")

            if unmatched_src:
                preview_text.insert(END, f"⚠️ 数据源中不存在 ({len(unmatched_src)}): {', '.join(unmatched_src)}\n")
            if unmatched_tgt:
                preview_text.insert(END, f"⚠️ 未找到模板 ({len(unmatched_tgt)}): {', '.join(unmatched_tgt)}\n")

            preview_text.config(state="disabled")

        ds_combo.bind("<<ComboboxSelected>>", update_preview)
        cfg_combo.bind("<<ComboboxSelected>>", update_preview)

        # 执行按钮
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=X, pady=5)

        def execute():
            result_text.config(state="normal")
            result_text.delete("1.0", END)
            result_text.update()

            ds = next((d for d in data_sources if d.name == ds_var.get()), None)
            t_item = next((t for t in templates if t.name == t_var.get()), None)
            cfg = next((c for c in configs if c.name == cfg_var.get()), None)

            if not all([ds, t_item, cfg]):
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

            try:
                source_wb = load_excel(ds.file_path, data_only=True)
                template_name_map = {os.path.basename(t.file_path): t for t in templates}
                target_path = t_item.file_path

                if is_file_locked(target_path):
                    target_path = _generate_copy_path(target_path)
                    shutil.copy2(t_item.file_path, target_path)
                    result_text.insert(END, f"📁 原始文件被占用，已生成副本\n\n")
                    result_text.update()

                target_wb = load_excel(target_path)
                copier = DataCopier(write_mode=WriteMode.OVERWRITE)
                all_results = []

                for mapping in cfg.config.sheet_mappings:
                    if mapping.data_source_sheet not in ds.sheets:
                        continue
                    if mapping.template_file not in template_name_map:
                        continue
                    result = copier.copy(source_wb, target_wb, mapping)
                    all_results.append(result)

                target_wb.save(target_path)

                success = [r for r in all_results if r.success]
                failed = [r for r in all_results if not r.success]

                result_text.insert(END, "="*40 + "\n")
                if success:
                    result_text.insert(END, f"✅ 成功: {len(success)} 个\n")
                    for r in success:
                        result_text.insert(END, f"  - {r.source_sheet} → {r.template_file}\n")
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
