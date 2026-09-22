import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import csv
import sys
from openpyxl import load_workbook

def resource_path(relative_path):
    # UNC网络共享路径直接返回，不拼接本地目录
    if relative_path.startswith(r"\\"):
        return relative_path
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ===== 替换原来的 json 加载方式 =====
from cable_list import CABLE_DATA
from board_list import BOARD_CFG
from flip_list import FLIP_LIST
from img_code_map import map_cable, map_board, map_conn, map_if, map_mode
# ==========【新增导入装配场景配置】==========
from scene_list import SCENE_LIST
# ==========【新增：软硬件通道映射配置】==========
from channel_map_list import CHANNEL_MAP_LIST


class PinConvertApp:
    def __init__(self, root):
        self.root = root
        self.root.title("线缆引脚定义生成工具v1.6")
        self.root.geometry("1080x520")

        # ========== 新增：顶部标签页容器 ==========
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=5)

        # 标签页1：引脚生成页面【增加横竖滚动，改为 self.tab1】
        self.tab1 = tk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="引脚定义生成")

        # ========== 【核心改动：tab1增加Canvas+横竖滚动条】 ==========
        tab1_canvas = tk.Canvas(self.tab1)
        scroll_v = ttk.Scrollbar(self.tab1, orient="vertical", command=tab1_canvas.yview)
        scroll_h = ttk.Scrollbar(self.tab1, orient="horizontal", command=tab1_canvas.xview)
        tab1_canvas.configure(yscrollcommand=scroll_v.set, xscrollcommand=scroll_h.set)

        tab1_canvas.grid(row=0, column=0, sticky="nsew")
        scroll_v.grid(row=0, column=1, sticky="ns")
        scroll_h.grid(row=1, column=0, sticky="ew")
        self.tab1.rowconfigure(0, weight=1)
        self.tab1.columnconfigure(0, weight=1)

        master_frame = tk.Frame(tab1_canvas)
        tab1_canvas.create_window((0,0), window=master_frame, anchor="nw")

        def update_scroll_region(event):
            tab1_canvas.configure(scrollregion=tab1_canvas.bbox("all"))
        master_frame.bind("<Configure>", update_scroll_region)

        # 鼠标滚轮：仅在tab1画布上悬浮才滚动（修复全局滚轮问题）
        def _on_mousewheel(event):
            if tab1_canvas.winfo_containing(event.x_root, event.y_root) == tab1_canvas:
                tab1_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        self.root.bind_all("<MouseWheel>", _on_mousewheel)
        # ========== 滚动区域改造完成 ==========

        master_frame.columnconfigure(0, weight=1)
        master_frame.columnconfigure(1, weight=0)

        # ====滚动信息内容====
        self.scroll_text = "欢迎使用线缆引脚定义生成工具    该工具仅限御渡半导体内部人员使用    如有优化建议或遇使用问题请联系开发人员    "
        self.scroll_pos = 0

        self.cable_data = CABLE_DATA
        self.board_cfg = BOARD_CFG
        self.flip_list = FLIP_LIST
        self.board_dict = self.board_cfg.get("boards", {})
        self.board_names = list(self.board_dict.keys())
        self.cable_list = list(self.cable_data.keys())
        self.line_widgets = []
        self.selected_row_idx = 0

        # ====第0行：线缆选择【增加线缆长度下拉 + 物料编码显示Label + 【查看线缆图】按钮】====
        row1 = tk.Frame(master_frame)
        row1.grid(row=0, column=0, columnspan=2, pady=20, padx=10, sticky="ew")
        tk.Label(row1, text="转接线缆选择：").pack(side="left")
        self.cable_combo = ttk.Combobox(row1, values=self.cable_list, state="readonly", width=20)
        self.cable_combo.pack(side="left", padx=8)
        self.cable_combo.bind("<<ComboboxSelected>>", self.on_cable_changed)
        if len(self.cable_list) > 0:
            self.cable_combo.current(0)

        # ====================【新增线缆长度下拉控件】====================
        tk.Label(row1, text="线缆长度：").pack(side="left", padx=(20, 0))
        self.cable_len_combo = ttk.Combobox(row1, state="readonly", width=10)
        self.cable_len_combo.pack(side="left", padx=8)
        self.cable_len_combo.bind("<<ComboboxSelected>>", self.on_length_changed)

        # ====================【新增：线缆物料编码只读显示框】====================
        tk.Label(row1, text="线缆物料编码：").pack(side="left", padx=(20, 0))
        self.cable_code_var = tk.StringVar()
        tk.Label(row1, textvariable=self.cable_code_var, width=12, bg="white", relief="solid").pack(side="left", padx=8)

        # ====================【新增：查看线缆图片按钮】====================
        self.cable_img_btn = ttk.Button(row1, text="查看线缆", command=self.show_cable_image)
        self.cable_img_btn.pack(side="left", padx=(10,0))

        # 第1行：左侧表格区域 + 右侧教程
        table_area = tk.Frame(master_frame)
        table_area.grid(row=1, column=0, padx=10, sticky="nw")
        self.grid_container = tk.Frame(table_area)
        self.grid_container.pack(anchor="nw")

        tutorial_frame = tk.LabelFrame(master_frame, text="使用教程", padx=8, pady=0)
        tutorial_frame.grid(row=1, column=1, padx=(0, 10), sticky="n")

        tutorial_text = (
            "1. 选择对应的转接线缆\n"
            "2. 逐行选择和线缆对接的板卡型号和连接器位号\n"
            "3. 选择线缆的对接方式（花面朝上/朝下）\n"
            "4. 点击【查看板卡图】查看当前行板卡示意图\n"
            "5. 点击【显示对接示意图】查看当前选中行的板卡连接器示意图\n"
            "6. 点击\"选择保存位置\"设置输出路径\n"
            "7. 点击\"生成引脚定义CSV文件\"\n"
            "8. CSV文件的G列为线缆输出端连接器PIN脚位号 \n"
            "9. CSV文件的C列为线缆输出端连接器PIN脚对应的网络名 \n"
        )
        tk.Label(tutorial_frame, text=tutorial_text, justify="left",
                 font=("微软雅黑", 9), fg="#333333").pack(anchor="w")

        # =========表格表头【新增一列：操作】=========
        tk.Label(self.grid_container, text="板卡型号", width=16).grid(row=0, column=0, padx=8, sticky="w")
        tk.Label(self.grid_container, text="板卡连接器位号", width=18).grid(row=0, column=1, padx=8, sticky="w")
        tk.Label(self.grid_container, text="线缆连接器位号", width=12).grid(row=0, column=2, padx=8, sticky="w")
        tk.Label(self.grid_container, text="对接方式", width=12).grid(row=0, column=3, padx=8, sticky="w")
        tk.Label(self.grid_container, text="操作", width=10).grid(row=0, column=4, padx=8, sticky="w")

        # ====【全局对接示意图按钮】跨两列，放在表格下方====
        img_btn_frame = tk.Frame(master_frame)
        img_btn_frame.grid(row=2, column=0, columnspan=1, pady=(0, 4), padx=10, sticky="w")
        self.global_img_btn = ttk.Button(img_btn_frame, text="显示对接示意图", command=self.show_selected_row_image)
        self.global_img_btn.pack(side="left")

        # ====保存路径输入行，跨左右两列====
        row_path = tk.Frame(master_frame)
        row_path.grid(row=3, column=0, columnspan=2, pady=(4, 8), padx=10, sticky="ew")

        tk.Label(row_path, text="输出文件路径：").pack(side="left")
        self.save_path_var = tk.StringVar()
        tk.Entry(row_path, textvariable=self.save_path_var, width=56).pack(side="left", padx=6)
        tk.Button(row_path, text="选择保存位置", command=self.select_save_path).pack(side="left")

        # ====【重点：生成CSV按钮位置保持不变，仍然在 row=4】====
        gen_btn = tk.Button(master_frame, text="生成转接后引脚定义CSV文件", command=self.generate_csv, bg="#337ab7", fg="white",
                            height=2)
        gen_btn.grid(row=4, column=0, columnspan=2, pady=20)

        # ====底部状态栏，放在master_frame最后一行，跨两列====
        status_bar = tk.Frame(master_frame, relief="sunken", bd=1, bg="#f0f0f0")
        status_bar.grid(row=5, column=0, columnspan=2, sticky="ew")

        tk.Label(status_bar, text="作者：冯志鹏", bg="#f0f0f0", font=("微软雅黑", 9)).pack(side="left", padx=10)
        self.scroll_label = tk.Label(status_bar, text="", bg="#f0f0f0", font=("微软雅黑", 9), fg="#333333")
        self.scroll_label.pack(side="left", fill="x", expand=True)
        tk.Label(status_bar, text="版本：v1.6", bg="#f0f0f0", font=("微软雅黑", 9)).pack(side="right", padx=10)

        self.update_scroll()
        self.on_cable_changed()

        # ========== 标签页2：接线装配说明==========
        self.tab2 = tk.Frame(self.notebook)
        self.notebook.add(self.tab2, text="接线装配说明")

        tab2_main = tk.Frame(self.tab2)
        tab2_main.pack(fill="both", expand=True, padx=20, pady=20)

        # 第一行：场景下拉 + 看图按钮
        scene_frame = tk.Frame(tab2_main)
        scene_frame.pack(anchor="w", pady=(0,10))
        tk.Label(scene_frame, text="装配场景选择：", font=("微软雅黑",11)).pack(side="left")
        self.scene_combo = ttk.Combobox(scene_frame, state="readonly", width=20)
        self.scene_combo.pack(side="left", padx=10)
        scene_btn = ttk.Button(scene_frame, text="查看装配示意图", command=self.show_scene_image)
        scene_btn.pack(side="left", padx=10)
        # 绑定场景切换事件，自动加载文档列表
        self.scene_combo.bind("<<ComboboxSelected>>", self.on_scene_selected)

        # 第二块：关联文档区域
        doc_frame = tk.LabelFrame(tab2_main, text="该场景关联文档", padx=10, pady=10)
        doc_frame.pack(fill="both", expand=True)

        # 文档列表 + 滚动条
        list_container = tk.Frame(doc_frame)
        list_container.pack(fill="both", expand=True)
        self.scene_doc_listbox = tk.Listbox(list_container, font=("微软雅黑", 9))
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=self.scene_doc_listbox.yview)
        self.scene_doc_listbox.config(yscrollcommand=scrollbar.set)
        self.scene_doc_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        # 绑定双击打开文档
        self.scene_doc_listbox.bind("<Double-Button-1>", lambda e: self.open_selected_scene_doc())

        # 文档操作按钮行
        doc_btn_frame = tk.Frame(doc_frame)
        doc_btn_frame.pack(pady=(8,0), anchor="w")
        self.open_doc_btn = ttk.Button(doc_btn_frame, text="打开选中文档", command=self.open_selected_scene_doc)
        self.open_doc_btn.pack()

        # 初始化场景下拉选项
        self.scene_cfg = SCENE_LIST
        scene_names = list(self.scene_cfg.keys())
        self.scene_combo["values"] = scene_names
        if scene_names:
            self.scene_combo.current(0)
            self.on_scene_selected()

        # ======================【标签页3：软硬件通道映射（动态读取Excel版）】======================
        self.tab3 = tk.Frame(self.notebook)
        self.notebook.add(self.tab3, text="软硬件通道映射")
        tab3_main = tk.Frame(self.tab3)
        tab3_main.pack(fill="both", expand=True, padx=30, pady=20)

        # 第1行：板卡型号下拉
        row_ch_board = tk.Frame(tab3_main)
        row_ch_board.pack(anchor="w", pady=(0, 8))
        tk.Label(row_ch_board, text="板卡型号：", font=("微软雅黑", 11)).pack(side="left")
        self.ch_board_combo = ttk.Combobox(row_ch_board, state="readonly", width=25)
        self.ch_board_combo.pack(side="left", padx=10)
        self.ch_board_combo.bind("<<ComboboxSelected>>", self.on_ch_board_selected)

        # 第2块：动态表单区域（列名+取值，动态生成）
        self.ch_form_frame = tk.LabelFrame(tab3_main, text="通道映射信息（对应Excel表格）", padx=12, pady=8)
        self.ch_form_frame.pack(fill="both", expand=True, pady=(4, 10), anchor="w")

        # 第3行：打开完整Excel映射文件按钮
        row_btn = tk.Frame(tab3_main)
        row_btn.pack(anchor="w", pady=(0, 0))
        self.open_ch_excel_btn = ttk.Button(row_btn, text="打开完整通道映射文件(Excel)", command=self.open_channel_map_file)
        self.open_ch_excel_btn.pack()

        # tab3 运行时状态
        self.ch_excel_rows = []        # Excel数据行（dict，key为列名）
        self.ch_column_names = []      # 配置的列名（显示顺序）
        self.ch_key_column = ""        # 关键列（下拉框列）
        self.ch_form_rows = {}         # 列名 -> 控件（下拉框/变量）
        self.ch_key_combo = None       # 关键列下拉框控件

        # 初始化板卡下拉
        self.channel_cfg = CHANNEL_MAP_LIST
        ch_board_names = list(self.channel_cfg.keys())
        self.ch_board_combo["values"] = ch_board_names
        if ch_board_names:
            self.ch_board_combo.current(0)
            self.on_ch_board_selected()

    # ============================================================
    # 标签页3：软硬件通道映射（动态读取Excel版）
    # ============================================================
    def _clear_ch_form(self):
        """清空通道映射表单，重置运行时状态。"""
        for w in self.ch_form_frame.winfo_children():
            w.destroy()
        self.ch_excel_rows = []
        self.ch_column_names = []
        self.ch_key_column = ""
        self.ch_form_rows = {}
        self.ch_key_combo = None

    def _show_form_status(self, msg):
        """在表单区域显示提示信息（错误/无配置等）。"""
        self._clear_ch_form()
        tk.Label(self.ch_form_frame, text=msg, fg="#a00",
                 font=("微软雅黑", 10), justify="left").grid(row=0, column=0, padx=6, pady=10, sticky="w")

    def _read_excel_rows(self, full_path):
        """
        读取通道映射Excel（openpyxl，无pandas依赖），按 self.ch_column_names 配置的列提取数据。
        自动定位包含全部配置列名的表头行，返回数据行列表（每行为 dict）。
        """
        wb = load_workbook(filename=full_path, data_only=True)
        ws = wb.active
        # 一次性读取所有行（普通模式，小文件无压力）
        all_rows = list(ws.iter_rows(values_only=True))
        wb.close()

        if not all_rows:
            raise Exception("Excel文件为空或没有读到任何行")

        # 1) 定位表头行：找到同时包含所有配置列名的那一行
        header_row_idx = None
        header = None
        for i, row in enumerate(all_rows):
            # 把行内所有单元格转成字符串集合（去空格、跳过None）
            cell_vals = set()
            for v in row:
                if v is not None and str(v).strip() != "":
                    cell_vals.add(str(v).strip())
            if all(col in cell_vals for col in self.ch_column_names):
                header_row_idx = i
                header = row
                break

        if header is None:
            # 没找到同时包含所有列名的行，列出Excel中实际读到的表头供排查
            sample_header = [str(v).strip() for v in all_rows[0] if v is not None] if all_rows else []
            raise Exception(
                f"Excel中找不到包含所有配置列名的表头行。\n"
                f"需要的列：{self.ch_column_names}\n"
                f"Excel首行读到：{sample_header}\n"
                f"请检查channel_map_list.py中columns列名是否和Excel表头完全一致"
            )

        # 2) 表头 -> 实际列索引
        header_str = [str(v).strip() if v is not None else "" for v in header]
        col_index_map = {}
        missing_cols = []
        for col_name in self.ch_column_names:
            if col_name in header_str:
                col_index_map[col_name] = header_str.index(col_name)
            else:
                missing_cols.append(col_name)
        if missing_cols:
            raise Exception(
                f"Excel表头缺少配置列：{missing_cols}\n"
                f"实际表头：{header_str}"
            )

        # 3) 从表头下一行开始逐行提取配置列的值
        rows = []
        for row in all_rows[header_row_idx + 1:]:
            rec = {}
            for col_name, idx in col_index_map.items():
                val = row[idx] if idx < len(row) else None
                if val is None or str(val).strip() == "":
                    rec[col_name] = ""
                else:
                    rec[col_name] = str(val).strip()
            # 跳过整行都为空的记录
            if any(rec.values()):
                rows.append(rec)
        return rows

    def on_ch_board_selected(self, event=None):
        """板卡型号切换：读取该板卡关联的Excel，并动态重建表单。"""
        self._clear_ch_form()
        board_name = self.ch_board_combo.get()
        if not board_name:
            return

        board_info = self.channel_cfg.get(board_name, {})
        # 取出列配置（去空白项）
        self.ch_column_names = [c for c in board_info.get("columns", []) if str(c).strip()]
        self.ch_key_column = board_info.get("key_column", "").strip()

        raw_path = board_info.get("full_map_file", "").strip()
        if not raw_path:
            self._show_form_status(f"板卡[{board_name}]未配置 full_map_file（Excel文件路径）")
            return

        full_path = resource_path(raw_path)
        if not os.path.exists(full_path):
            is_unc = full_path.startswith(r"\\")
            tip = "\n请确认已连接共享盘、有访问权限" if is_unc else ""
            self._show_form_status(f"找不到通道映射Excel：\n{full_path}{tip}")
            return

        if not self.ch_column_names:
            self._show_form_status(f"板卡[{board_name}]未配置 columns（需显示的列名）")
            return

        try:
            self.ch_excel_rows = self._read_excel_rows(full_path)
        except Exception as e:
            self._show_form_status(f"读取Excel失败：\n{e}")
            return

        if not self.ch_excel_rows:
            self._show_form_status(f"Excel中未找到可显示的数据行，请检查列名配置：\n{self.ch_column_names}")
            return

        self._build_ch_form()

    def _build_ch_form(self):
        """按配置列动态生成表单：左侧列名Label，右侧取值控件（关键列为下拉框，其余为只读框）。"""
        if not self.ch_column_names:
            return

        for i, col in enumerate(self.ch_column_names):
            # 左侧：列名文本
            tk.Label(self.ch_form_frame, text=col, font=("微软雅黑", 11)).grid(
                row=i, column=0, sticky="w", padx=6, pady=6)

            if col == self.ch_key_column:
                # 关键列（嵌入式通道）：下拉框，选项取Excel该列的全部唯一值
                values = list(dict.fromkeys(
                    r.get(self.ch_key_column, "") for r in self.ch_excel_rows
                    if r.get(self.ch_key_column, "")
                ))
                combo = ttk.Combobox(self.ch_form_frame, state="readonly", width=30)
                combo["values"] = values
                combo.grid(row=i, column=1, sticky="w", padx=6, pady=6)
                combo.bind("<<ComboboxSelected>>", self.on_ch_selected)
                self.ch_key_combo = combo
                self.ch_form_rows[col] = combo
            else:
                # 其余列：只读显示框，选中关键列后自动填充对应值
                var = tk.StringVar()
                lbl = tk.Label(self.ch_form_frame, textvariable=var, width=30, bg="white",
                               relief="solid", anchor="w", font=("微软雅黑", 11))
                lbl.grid(row=i, column=1, sticky="w", padx=6, pady=6)
                self.ch_form_rows[col] = var

        # 默认选中关键列第一项，联动显示其余列
        if self.ch_key_combo is not None and self.ch_key_combo["values"]:
            self.ch_key_combo.current(0)
            self.on_ch_selected()

    def on_ch_selected(self, event=None):
        """关键列（嵌入式通道）选中：按Excel对应行刷新其余列取值。"""
        if self.ch_key_combo is None or not self.ch_key_column:
            return
        key_val = self.ch_key_combo.get()

        # 在Excel数据中查找关键列值匹配的那一行
        matched = None
        for r in self.ch_excel_rows:
            if str(r.get(self.ch_key_column, "")) == key_val:
                matched = r
                break

        for col, widget in self.ch_form_rows.items():
            if col == self.ch_key_column:
                continue
            val = matched.get(col, "") if matched else ""
            widget.set(str(val) if val not in (None, "nan") else "")

    # ========== 标签页3回调：打开当前板卡对应的Excel映射文件 ==========
    def open_channel_map_file(self):
        board_name = self.ch_board_combo.get()
        board_info = self.channel_cfg.get(board_name, {})
        raw_path = board_info.get("full_map_file", "").strip()
        if not raw_path:
            messagebox.showinfo("提示", "当前板卡未配置完整映射文件路径")
            return
        full_path = resource_path(raw_path)
        is_unc = full_path.startswith(r"\\")
        if not os.path.exists(full_path):
            tip = ""
            if is_unc:
                tip = "\n⚠️网络共享路径，请确认已连接共享盘、有访问权限"
            messagebox.showerror("文件不存在", f"找不到通道映射Excel：\n{full_path}{tip}")
            return
        try:
            os.startfile(full_path)
        except Exception as e:
            err_msg = str(e)
            if is_unc:
                err_msg += "\n可能原因：未映射共享盘/无访问权限/共享断开"
            messagebox.showerror("打开失败", f"无法打开Excel：{err_msg}")

    # ==========【tab2新增：场景切换回调，刷新文档列表】==========
    def on_scene_selected(self, event=None):
        scene_name = self.scene_combo.get()
        scene_info = self.scene_cfg.get(scene_name, {})
        doc_list = scene_info.get("docs", [])
        self.scene_doc_listbox.delete(0, tk.END)
        # 存入实例变量保存文档路径映射
        self._scene_docs_cache = doc_list
        for doc_item in doc_list:
            display_name = doc_item.get("name", "")
            self.scene_doc_listbox.insert(tk.END, display_name)

    # ==========【tab2新增：打开选中文档，支持UNC网络共享路径】==========
    def open_selected_scene_doc(self):
        sel_idx = self.scene_doc_listbox.curselection()
        if not sel_idx:
            messagebox.showinfo("提示", "请先在列表选中一个文档")
            return
        idx = sel_idx[0]
        doc_item = self._scene_docs_cache[idx]
        raw_path = doc_item.get("path", "").strip()
        full_path = resource_path(raw_path)

        is_unc = full_path.startswith(r"\\")
        if not os.path.exists(full_path):
            tip = ""
            if is_unc:
                tip = "\n⚠️网络共享路径，请确认已连接共享盘、有访问权限"
            messagebox.showerror("文件不存在", f"找不到文档：\n{full_path}{tip}")
            return
        try:
            os.startfile(full_path)
        except Exception as e:
            err_msg = str(e)
            if is_unc:
                err_msg += "\n可能原因：未映射共享盘/无访问权限/共享断开"
            messagebox.showerror("打开失败", f"无法打开文档：{err_msg}")

    # ==========【新增函数：打开线缆图片弹窗】==========
    def show_cable_image(self):
        cable_name = self.cable_combo.get()
        base_name = f"cable_{cable_name}"
        img_folder = resource_path("conn_img")
        png_path = os.path.join(img_folder, f"{base_name}.png")
        gif_path = os.path.join(img_folder, f"{base_name}.gif")
        self.show_full_conn_image(base_name, png_path, gif_path)

    # ==========【新增函数：打开板卡图片弹窗】==========
    def show_board_image(self, board_name):
        base_name = f"board_{board_name}"
        img_folder = resource_path("conn_img")
        png_path = os.path.join(img_folder, f"{base_name}.png")
        gif_path = os.path.join(img_folder, f"{base_name}.gif")
        self.show_full_conn_image(base_name, png_path, gif_path)

    # ==========【新增函数：打开装配场景图片弹窗】==========
    def show_scene_image(self):
        selected_scene = self.scene_combo.get()
        base_name = f"scene_{selected_scene}"
        img_folder = resource_path("conn_img")
        png_path = os.path.join(img_folder, f"{base_name}.png")
        gif_path = os.path.join(img_folder, f"{base_name}.gif")
        self.show_full_conn_image(base_name, png_path, gif_path)

    # ==========【新增函数：长度下拉切换，更新物料编码】==========
    def on_length_changed(self, event=None):
        cable_name = self.cable_combo.get()
        cfg = self.cable_data.get(cable_name, {})
        len_dict = cfg.get("length_list", {})
        selected_len = self.cable_len_combo.get()
        code = len_dict.get(selected_len, "")
        self.cable_code_var.set(code)

    def on_board_selected(self, board_combo, conn_combo, row_idx):
        self.selected_row_idx = row_idx
        selected_board = board_combo.get()
        conn_dict = self.board_dict.get(selected_board, {})
        conn_list = list(conn_dict.keys())
        conn_combo["values"] = conn_list
        if len(conn_list) > 0:
            conn_combo.current(0)
        else:
            conn_combo.set("")

    def on_mode_changed(self, row_idx, event):
        self.selected_row_idx = row_idx

    def on_cable_changed(self, event=None):
        cable_name = self.cable_combo.get()
        if cable_name not in self.cable_data:
            return
        cfg = self.cable_data[cable_name]
        input_list = cfg["input_if"]

        len_dict = cfg.get("length_list", {})
        len_options = list(len_dict.keys())
        default_len = cfg.get("default_len", "")
        self.cable_len_combo["values"] = len_options
        if len_options:
            self.cable_len_combo.set(default_len)
        else:
            self.cable_len_combo.set("")
        self.on_length_changed()

        for w in self.grid_container.winfo_children():
            w.destroy()
        self.line_widgets.clear()
        self.selected_row_idx = 0

        # 重建表头
        tk.Label(self.grid_container, text="板卡型号", width=16).grid(row=0, column=0, padx=8, sticky="w")
        tk.Label(self.grid_container, text="板卡连接器位号", width=18).grid(row=0, column=1, padx=8, sticky="w")
        tk.Label(self.grid_container, text="线缆连接器位号", width=12).grid(row=0, column=2, padx=8, sticky="w")
        tk.Label(self.grid_container, text="对接方式", width=12).grid(row=0, column=3, padx=8, sticky="w")
        tk.Label(self.grid_container, text="操作", width=10).grid(row=0, column=4, padx=8, sticky="w")

        for row_idx, item in enumerate(input_list, start=1):
            cb_board = ttk.Combobox(self.grid_container, values=self.board_names, state="readonly", width=16)
            if len(self.board_names) > 0:
                cb_board.current(0)
            cb_board.grid(row=row_idx, column=0, padx=8, pady=3, sticky="w")

            cb_conn = ttk.Combobox(self.grid_container, values=[], state="readonly", width=22)
            cb_conn.grid(row=row_idx, column=1, padx=8, pady=3, sticky="w")

            lbl_if = tk.Label(self.grid_container, text=item, width=12)
            lbl_if.grid(row=row_idx, column=2, padx=8, pady=3, sticky="w")

            cb_mode = ttk.Combobox(self.grid_container, values=["花面朝上", "花面朝下"], state="readonly", width=12)
            cb_mode.current(0)
            cb_mode.grid(row=row_idx, column=3, padx=8, pady=3, sticky="w")

            # ====================【新增：板卡图片按钮】====================
            btn_board_img = ttk.Button(self.grid_container, text="查看板卡", width=8,
                                       command=lambda b=cb_board: self.show_board_image(b.get()))
            btn_board_img.grid(row=row_idx, column=4, padx=8, pady=3)

            cb_board.bind("<<ComboboxSelected>>",
                          lambda e, b=cb_board, c=cb_conn, r=row_idx: self.on_board_selected(b, c, r))
            cb_mode.bind("<<ComboboxSelected>>", lambda e, r=row_idx: self.on_mode_changed(r, e))

            self.line_widgets.append({
                "if_name": item,
                "board_combo": cb_board,
                "conn_combo": cb_conn,
                "mode_combo": cb_mode,
            })
            self.on_board_selected(cb_board, cb_conn, row_idx)

    def show_selected_row_image(self):
        if len(self.line_widgets) == 0:
            messagebox.showwarning("提示", "当前线缆没有输入接口！")
            return

        cable_name = self.cable_combo.get()
        cable_id = map_cable.get(cable_name, 99)
        parts_num = [str(cable_id)]

        for line in self.line_widgets:
            board_name = line["board_combo"].get()
            conn_name = line["conn_combo"].get()
            wire_if = line["if_name"]
            mode = line["mode_combo"].get()

            b_id = map_board.get(board_name, 99)
            c_id = 99
            for key, val in map_conn.items():
                if key in conn_name:
                    c_id = val
                    break
            i_id = map_if.get(wire_if, 99)
            m_id = map_mode.get(mode, 99)

            seg_num = f"{b_id}{c_id}{i_id}{m_id}"
            parts_num.append(seg_num)

        base_name = "_".join(parts_num)
        img_folder = resource_path("conn_img")
        png_path = os.path.join(img_folder, f"{base_name}.png")
        gif_path = os.path.join(img_folder, f"{base_name}.gif")

        self.show_full_conn_image(base_name, png_path, gif_path)

    def show_full_conn_image(self, base_name, png_path, gif_path):
        img_win = tk.Toplevel(self.root)
        img_win.title(f"示意图 - {base_name}")
        img_win.after_id = None

        if os.path.exists(gif_path):
            img_win.frames = []
            img_win.frame_index = 0
            try:
                frame_num = 0
                while True:
                    frame = tk.PhotoImage(file=gif_path, format=f"gif -index {frame_num}")
                    img_win.frames.append(frame)
                    frame_num += 1
            except tk.TclError:
                pass

            if len(img_win.frames) == 0:
                messagebox.showerror("GIF加载失败", "GIF无有效帧！")
                img_win.destroy()
                return

            label_img = tk.Label(img_win)
            label_img.pack(expand=True)

            def animate():
                if not img_win.winfo_exists():
                    return
                frame = img_win.frames[img_win.frame_index]
                label_img.config(image=frame)
                img_win.frame_index += 1
                if img_win.frame_index >= len(img_win.frames):
                    img_win.frame_index = 0
                img_win.after_id = img_win.after(150, animate)

            animate()
        elif os.path.exists(png_path):
            try:
                photo = tk.PhotoImage(file=png_path)
                img_win.photo = photo
                label_img = tk.Label(img_win, image=photo)
                label_img.pack(expand=True)
            except Exception as e:
                messagebox.showerror("PNG加载失败", f"{e}")
                img_win.destroy()
                return
        else:
            messagebox.showerror("图片不存在",
                                 f"找不到示意图文件：\n{base_name}.gif 或 {base_name}.png\n请检查接线方式是否正确")
            img_win.destroy()
            return

        def on_close():
            if img_win.after_id is not None:
                img_win.after_cancel(img_win.after_id)
            img_win.destroy()

        img_win.protocol("WM_DELETE_WINDOW", on_close)

    def select_save_path(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")],
            title="保存引脚定义文件"
        )
        if path:
            self.save_path_var.set(path)

    def generate_csv(self):
        cable_name = self.cable_combo.get()
        cable_length = self.cable_len_combo.get()
        cable_code = self.cable_code_var.get()
        out_path = self.save_path_var.get().strip()
        if not out_path:
            messagebox.showwarning("提示", "请先选择输出路径！")
            return

        cfg = self.cable_data[cable_name]
        pin_map = cfg["pin_mapping"]
        pin_list = list(pin_map.items())

        line_setting = {}
        for line in self.line_widgets:
            line_setting[line["if_name"]] = {
                "board": line["board_combo"].get(),
                "connector": line["conn_combo"].get(),
                "mode": line["mode_combo"].get()
            }

        headers = ["板卡型号", "板卡连接器位号", "网络", "线缆连接器位号", "输入PIN脚", "对接方式", "输出PIN脚", "线缆型号", "线缆长度", "线缆物料编码"]
        rows = []

        for dest_pin, src_pin in pin_list:
            if not src_pin:
                if_name = ""
                pin_num = ""
            elif "_" in src_pin:
                idx = src_pin.rfind("_")
                if_name = src_pin[:idx]
                pin_num = src_pin[idx + 1:]
            else:
                if_name = ""
                pin_num = ""

            s = line_setting.get(if_name, {"board": "", "connector": "", "mode": ""})
            lookup_pin = pin_num

            conn_name = s["connector"]
            gui_mode = s["mode"]
            # ===== 你原来后面剩下的生成逻辑继续在这里补全 =====

    def update_scroll(self):
        self.scroll_label.config(text=self.scroll_text[self.scroll_pos:] + self.scroll_text[:self.scroll_pos])
        self.scroll_pos += 1
        if self.scroll_pos >= len(self.scroll_text):
            self.scroll_pos = 0
        self.root.after(200, self.update_scroll)


if __name__ == "__main__":
    root = tk.Tk()
    app = PinConvertApp(root)
    root.mainloop()
