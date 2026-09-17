import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import csv
import sys

def resource_path(relative_path):
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

class PinConvertApp:
    def __init__(self, root):
        self.root = root
        self.root.title("线缆引脚定义生成工具v1.4")
        self.root.geometry("1080x420")

        # ========== 新增：顶部标签页容器 ==========
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=5)

        # 标签页1：原有引脚生成页面
        self.tab1 = tk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="引脚定义生成")

        # 标签页2：新建页面【接线装配说明】
        self.tab2 = tk.Frame(self.notebook)
        self.notebook.add(self.tab2, text="接线装配说明")

        # ====================== 把原来master_frame全部放到 tab1 里面 ======================
        master_frame = tk.Frame(self.tab1)
        master_frame.pack(fill="both", expand=True)

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
        gen_btn.grid(row=3, column=1, columnspan=2, pady=20)

        # ====底部状态栏，放在master_frame最后一行，跨两列====
        status_bar = tk.Frame(master_frame, relief="sunken", bd=1, bg="#f0f0f0")
        status_bar.grid(row=5, column=0, columnspan=2, sticky="ew")

        tk.Label(status_bar, text="作者：冯志鹏", bg="#f0f0f0", font=("微软雅黑", 9)).pack(side="left", padx=10)
        self.scroll_label = tk.Label(status_bar, text="", bg="#f0f0f0", font=("微软雅黑", 9), fg="#333333")
        self.scroll_label.pack(side="left", fill="x", expand=True)
        tk.Label(status_bar, text="版本：v1.4", bg="#f0f0f0", font=("微软雅黑", 9)).pack(side="right", padx=10)

        self.update_scroll()
        self.on_cable_changed()

        # ========== 【新页面tab2 接线装配说明 UI】==========
        scene_frame = tk.Frame(self.tab2)
        scene_frame.pack(pady=40, padx=20, anchor="w")

        tk.Label(scene_frame, text="装配场景选择：", font=("微软雅黑",11)).pack(side="left")
        # ==========【修改：从导入的SCENE_LIST读取，不再硬编码】==========
        self.scene_list = SCENE_LIST
        self.scene_combo = ttk.Combobox(scene_frame, values=self.scene_list, state="readonly", width=20)
        self.scene_combo.pack(side="left", padx=10)
        if self.scene_list:
            self.scene_combo.current(0)

        scene_btn = ttk.Button(scene_frame, text="查看装配示意图", command=self.show_scene_image)
        scene_btn.pack(side="left", padx=10)


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
                                 f"找不到示意图文件：\n{base_name}.gif 或 {base_name}.png\n请检查conn_img文件夹")
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
            # =========修复翻转判断逻辑=========
            if conn_name and gui_mode == "花面朝下":
                flip_dict = self.flip_list.get("CONNECTOR_C", {})
                lookup_pin = flip_dict.get(lookup_pin, lookup_pin)

            net_name = self.board_dict.get(s["board"], {}).get(conn_name, {}).get(lookup_pin, "")
            row = [s["board"], conn_name, net_name, if_name, src_pin, s["mode"], dest_pin, cable_name, cable_length,
                   cable_code]
            rows.append(row)

        try:
            with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                writer.writerows(rows)
            messagebox.showinfo("成功", "CSV文件生成完成")
        except Exception as e:
            messagebox.showerror("保存失败", str(e))

    def update_scroll(self):
        self.scroll_pos += 1
        if self.scroll_pos >= len(self.scroll_text):
            self.scroll_pos = 0
        self.scroll_label.config(text=self.scroll_text[self.scroll_pos:] + self.scroll_text[:self.scroll_pos])
        self.root.after(120, self.update_scroll)

if __name__ == "__main__":
    root = tk.Tk()
    app = PinConvertApp(root)
    root.mainloop()
