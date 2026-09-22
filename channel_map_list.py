# -*- coding: utf-8 -*-
"""
软硬件通道映射配置
====================
本文件为“软硬件通道映射”标签页的数据源，在程序启动时由主程序导入：

  - 板卡型号：CHANNEL_MAP_LIST 的 key（第一行下拉框选项）
  - full_map_file：该板卡关联的完整通道映射Excel文件（选择板卡后读取）
  - columns：需要从Excel中提取并显示的列名（顺序即界面显示顺序）
            其中“嵌入式通道”等列名必须与Excel表头完全一致
  - key_column：作为下拉框的关键列（即“嵌入式通道”列）
            选中该列某一行值后，界面其余列自动显示Excel中对应行的内容

说明：
  1. 请按实际板卡型号、列名替换下方【示例数据】。
  2. 每种板卡型号拥有独立的分位Excel文件和独立的列配置。
  3. full_map_file 支持两种路径：
       - 相对路径：相对程序目录，如 r"channel_map\主板A_通道映射.xlsx"
       - UNC网络共享路径：以两个反斜杠开头，如 r"\\\\server\\share\\主板A_通道映射.xlsx"
  4. 完整通道映射Excel文件请提前放入程序目录下的 channel_map 文件夹，
     文件名与下方 full_map_file 保持一致。
"""

CHANNEL_MAP_LIST = {
    # ================= 示例板卡A =================
    "数字板D800": {
        # 该板卡完整的通道映射Excel文件（提前放入 channel_map 文件夹）
        "full_map_file": r"channel_map\数字板D800_软硬件通道映射.xlsx",
        # 需要从Excel中提取并显示的列（顺序即界面显示顺序）
        "columns": [
            "嵌入式对应通道",   # 下拉框列（key_column）
            "DUT网络名",
            "对应MVP",
            "PE位号",
        ],
        # 作为下拉框的关键列，必须存在于 columns 中
        "key_column": "嵌入式对应通道",
    },

    # ================= 示例板卡B =================
    "DEMO_BOARD_B": {
        "full_map_file": r"channel_map\DEMO_BOARD_B_通道映射.xlsx",
        "columns": [
            "嵌入式通道",
            "DUT网络名",
            "信号类型",
        ],
        "key_column": "嵌入式通道",
    },
}
