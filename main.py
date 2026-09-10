import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
import csv

# ===== 替换原来的 json 加载方式 =====
from cable_list import CABLE_DATA
from board_list import BOARD_CFG
from flip_list import FLIP_LIST

# CABLE_CONFIG = "cable_list.json"
# BOARD_CONFIG = "board_list.json"
# FLIP_LIST = "flip_list.json"

class PinConvertApp:
    def __init__(self, root):

        self.root = root
        self.root.title("线缆引脚定义生成工具v1.1")
        self.root.geometry("1000x450")
        # ====滚动信息内容====

        self.scroll_text = "欢迎使用线缆引脚定义生成工具    该工具仅限御渡半导体内部人员使用    如有优化建议或遇使用问题请联系开发人员    "
        self.scroll_pos = 0

        # 直接使用导入的数据，不再读文件
        self.cable_data = CABLE_DATA
        self.board_cfg = BOARD_CFG
        self.flip_list = FLIP_LIST
        self.board_dict = self.board_cfg.get("boards", {})
        self.board_names = list(self.board_dict.keys())
        self.cable_list = list(self.cable_data.keys())
        self.line_widgets = []

        # self.cable_data = self.load_json(CABLE_CONFIG)
        # self.board_cfg = self.load_json(BOARD_CONFIG)
        # self.flip_list = self.load_json(FLIP_LIST)
        # self.board_dict = self.board_cfg.get("boards", {})
        # self.board_names = list(self.board_dict.keys())
        # self.cable_list = list(self.cable_data.keys()) #修复缺失行
        # self.line_widgets = []

        #====第1行：线缆选择====
        row1 = tk.Frame(root)
        row1.pack(pady=6, fill="x", padx=10)
        tk.Label(row1, text="转接线缆选择：").pack(side="left")
        self.cable_combo = ttk.Combobox(row1, values=self.cable_list, state="readonly", width=30)
        self.cable_combo.pack(side="left", padx=8)
        self.cable_combo.bind("<<ComboboxSelected>>", self.on_cable_changed)
        if len(self.cable_list) > 0:
            self.cable_combo.current(0)


        # #====标题行，改用grid实现严格对齐====
        # self.grid_container = tk.Frame(root)
        # self.grid_container.pack(pady=2, padx=10, anchor="w") #容器整体靠左
        # ====主区域：左侧表格 + 右侧教程====
        main_area = tk.Frame(root)
        main_area.pack(pady=2, padx=10, fill="x")

        # 左侧：表格容器
        self.grid_container = tk.Frame(main_area)
        self.grid_container.pack(side="left", anchor="w")

        # 右侧：使用教程
        tutorial_frame = tk.LabelFrame(main_area, text="使用教程", padx=8, pady=6)
        tutorial_frame.pack(side="right", padx=20, anchor="n")

        tutorial_text = (
            "1. 选择对应的转接线缆\n"
            "2. 逐行选择和线缆对接的板卡型号和连接器位号\n"
            "3. 选择线缆的对接方式（花面朝上/朝下）\n"
            "4. 点击\"选择保存位置\"设置输出路径\n"
            "5. 点击\"生成引脚定义CSV文件\"\n"
            "6. CSV文件的G列为线缆输出端连接器PIN脚位号 \n"
            "7. CSV文件的C列为线缆输出端连接器PIN脚对应的网络名 \n"
            "\n"
            "注意事项：\n"
            "- 对接方式需与实际接线方式一致\n"
            "- CSV文件可用Excel直接打开\n"
            "- 如有问题请联系管理员"
        )
        tk.Label(tutorial_frame, text=tutorial_text, justify="left",
                 font=("微软雅黑", 9), fg="#333333").pack(anchor="w")

        tk.Label(self.grid_container, text="板卡型号", width=16).grid(row=0, column=0, padx=8, sticky="w")
        tk.Label(self.grid_container, text="板卡连接器位号", width=18).grid(row=0, column=1, padx=8, sticky="w")
        tk.Label(self.grid_container, text="线缆连接器位号", width=12).grid(row=0, column=2, padx=8, sticky="w")
        tk.Label(self.grid_container, text="对接方式", width=12).grid(row=0, column=3, padx=8, sticky="w")



        #====保存路径====
        row_path = tk.Frame(root)
        row_path.pack(pady=8, fill="x", padx=10)
        tk.Label(row_path, text="输出文件路径：").pack(side="left")
        self.save_path_var = tk.StringVar()
        tk.Entry(row_path, textvariable=self.save_path_var, width=56).pack(side="left", padx=6)
        tk.Button(row_path, text="选择保存位置", command=self.select_save_path).pack(side="left")

        #====生成按钮====
        tk.Button(root, text="生成引脚定义CSV文件", command=self.generate_csv, bg="#337ab7", fg="white",
                  height=2).pack(pady=16)

        # ====底部状态栏====
        status_bar = tk.Frame(root, relief="sunken", bd=1, bg="#f0f0f0")
        status_bar.pack(side="bottom", fill="x")

        # 左侧：作者
        tk.Label(status_bar, text="作者：冯志鹏", bg="#f0f0f0", font=("微软雅黑", 9)).pack(side="left", padx=10)

        # 中间：滚动信息
        self.scroll_label = tk.Label(status_bar, text="", bg="#f0f0f0", font=("微软雅黑", 9), fg="#333333")
        self.scroll_label.pack(side="left", fill="x", expand=True)

        # 右侧：版本号
        tk.Label(status_bar, text="版本：v1.1", bg="#f0f0f0", font=("微软雅黑", 9)).pack(side="right", padx=10)

        # 启动滚动
        self.update_scroll()

        self.on_cable_changed()

    def load_json(self, filename):
        if not os.path.exists(filename):
            messagebox.showerror("错误", "找不到配置文件 " + filename)
            return {}
        try:
            with open(filename, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            messagebox.showerror("读取配置失败", str(e))
            return {}

    def on_board_selected(self, board_combo, conn_combo):
        """板卡切换刷新连接器下拉"""
        selected_board = board_combo.get()
        conn_dict = self.board_dict.get(selected_board, {})
        conn_list = list(conn_dict.keys())
        conn_combo["values"] = conn_list
        if len(conn_list) > 0:
            conn_combo.current(0)
        else:
            conn_combo.set("")

    def on_cable_changed(self, event=None):
        cable_name = self.cable_combo.get()
        if cable_name not in self.cable_data:
            return
        cfg = self.cable_data[cable_name]
        input_list = cfg["input_if"]

        #清除所有数据行，保留第0行标题不动
        for w in self.grid_container.winfo_children():
            grid_info = w.grid_info()
            if int(grid_info["row"]) > 0:
                w.destroy()
        self.line_widgets.clear()

        #从第1行开始填充数据
        for row_idx, item in enumerate(input_list, start=1):
            #第0列：板卡型号下拉框
            cb_board = ttk.Combobox(self.grid_container, values=self.board_names, state="readonly", width=16)
            if len(self.board_names) > 0:
                cb_board.current(0)
            cb_board.grid(row=row_idx, column=0, padx=8, pady=3, sticky="w")

            #第1列：连接器位号下拉框
            cb_conn = ttk.Combobox(self.grid_container, values=[], state="readonly", width=18)
            cb_conn.grid(row=row_idx, column=1, padx=8, pady=3, sticky="w")

            #第2列：输入接口文本
            tk.Label(self.grid_container, text=item, width=12).grid(row=row_idx, column=2, padx=8, pady=3, sticky="w")

            #第3列：接线方式下拉框
            cb_mode = ttk.Combobox(self.grid_container, values=["花面朝上", "花面朝下"], state="readonly", width=12)
            cb_mode.current(0)
            cb_mode.grid(row=row_idx, column=3, padx=8, pady=3, sticky="w")

            cb_board.bind("<<ComboboxSelected>>",
                          lambda e, b=cb_board, c=cb_conn: self.on_board_selected(b, c))

            self.line_widgets.append({
                "if_name": item,
                "board_combo": cb_board,
                "conn_combo": cb_conn,
                "mode_combo": cb_mode
            })
            self.on_board_selected(cb_board, cb_conn)

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
        out_path = self.save_path_var.get().strip()
        if not out_path:
            messagebox.showwarning("提示", "请先选择输出路径！")
            return

        cfg = self.cable_data[cable_name]
        pin_map = cfg["pin_mapping"]
        pin_list = list(pin_map.items())

        # 收集GUI界面上A1/A2选择的板卡、连接器、接线方式（原有逻辑不变）
        line_setting = {}
        for line in self.line_widgets:
            line_setting[line["if_name"]] = {
                "board": line["board_combo"].get(),
                "connector": line["conn_combo"].get(),
                "mode": line["mode_combo"].get()
            }

        headers = ["板卡型号", "板卡连接器位号", "网络", "线缆连接器位号", "输入PIN脚", "对接方式", "输出PIN脚", "线缆型号"]
        rows = []

        # dest_pin=输出引脚(B1_A1)，src_pin=输入引脚(A1_B8)
        for dest_pin, src_pin in pin_list:
            if not src_pin:  # value为空 ""
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

            # 判断花面朝向不一致，执行引脚映射
            conn_name = s["connector"]
            gui_mode = s["mode"]
            if conn_name and gui_mode and (gui_mode not in conn_name):
                # 读取全局pin_mapping总表，和json结构匹配
                flip_dict = self.flip_list.get("CONNECTOR_C", {})
                lookup_pin = flip_dict.get(lookup_pin, lookup_pin)

            # 根据【板卡+连接器+引脚号】查找网络名（使用转换后的lookup_pin）
            net_name = self.board_dict.get(s["board"], {}).get(conn_name, {}).get(lookup_pin, "")
            row = [s["board"], conn_name, net_name, if_name, src_pin, s["mode"], dest_pin, cable_name]
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
        """滚动信息更新"""
        self.scroll_pos += 1
        if self.scroll_pos >= len(self.scroll_text):
            self.scroll_pos = 0
        # 截取当前位置的文本，实现从右往左滚动效果
        display = self.scroll_text[self.scroll_pos:] + self.scroll_text[:self.scroll_pos]
        self.scroll_label.config(text=display)
        self.root.after(300, self.update_scroll)  # 每150ms刷新一次，数字越小滚动越快

if __name__ == "__main__":
    win = tk.Tk()
    app = PinConvertApp(win)
    win.mainloop()
