# ui_modules.py 开头的导入区（最终版）
import tkinter as tk
import os
from tkinter import ttk, messagebox
import pandas as pd

# 新增：导入拆分后的分析结果页模块（使用优化后的Treeview版本）
from ui_analysis import build_analysis_tab

# 仅导入手动导入函数 + 年预计量函数
from data_handling import (
    import_raw_material_manual,
    import_weekly_order_manual,
    import_daily_order_manual,  # 新增：每日订单导入
    import_yearly_estimate_manual,
    import_ordered_file_manual,  # 新增：已下单文件导入
    import_finished_stock_manual,
    import_raw_stock_manual,
    quick_import_all_last_files,  # 新增：快速导入所有上次文件的函数
    save_analysis_report, load_history_reports
)
from analysis_engine import run_analysis

# 全局变量（存储所有数据）
global_data = {
    "raw_material_df": None,    # 毛坯基础信息
    "weekly_order_df": None,    # 每周订单
    "yearly_estimate_df": None, # 年预计量
    "ordered_file_df": None,    # 新增：已下单文件
    "finished_stock_df": None,  # 当日成品库存
    "raw_stock_df": None,       # 当日毛坯库存
    "analysis_result_df": None, # 分析结果
    "analysis_show_df": None,   # 精简后的展示用分析结果
    "analysis_widgets": None,   # 分析结果页控件
    "last_import_paths": {}     # 存储上次导入的文件路径
}

def build_main_interface(root, global_data, path_config):
    """构建主界面：标签页 + 状态栏"""
    # 底部状态栏
    status_var = tk.StringVar(value="就绪 - 当前路径：{}".format(os.getcwd()))
    status_bar = ttk.Label(root, textvariable=status_var, relief=tk.SUNKEN)
    status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    # 创建标签页控件
    tab_control = ttk.Notebook(root)

    # 2. 数据导入标签页（后构建）
    tab_import = ttk.Frame(tab_control)
    tab_control.add(tab_import, text="数据导入")
    build_import_tab(tab_import, global_data, path_config, status_var)
    
    # 1. 分析结果标签页（先构建，调用拆分后的模块）
    tab_analysis = ttk.Frame(tab_control)
    tab_control.add(tab_analysis, text="分析结果")
    analysis_widgets = build_analysis_tab(tab_analysis, global_data, path_config)
    # 将analysis_widgets存入全局数据
    global_data["analysis_widgets"] = analysis_widgets
    
    # 3. 订单追踪标签页
    tab_maintain = ttk.Frame(tab_control)
    tab_control.add(tab_maintain, text="订单追踪")
    build_maintain_tab(tab_maintain, global_data)
    
    tab_control.pack(expand=1, fill="both")

def build_import_tab(parent, global_data, path_config, status_var):
    """构建数据导入标签页：新增已下单文件导入"""
    # 左侧操作区主容器（全程用pack）
    frame_left = ttk.Frame(parent, width=250)
    frame_left.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)
    
    # ========== 上方导入按钮分组容器 ==========
    frame_import_buttons = ttk.Frame(frame_left)
    frame_import_buttons.pack(side=tk.TOP, fill=tk.X)
    
    # 1. 毛坯基础信息导入
    ttk.Label(frame_import_buttons, text="毛坯基础信息导入", font=("微软雅黑", 12, "bold")).pack(pady=5, anchor="w")
    ttk.Button(frame_import_buttons, text="手动选择文件", 
               command=lambda: import_raw_material_manual(global_data, status_var)).pack(fill=tk.X, pady=3)
    
    # 2. 每周订单导入
    ttk.Label(frame_import_buttons, text="每周订单导入", font=('微软雅黑', 12, 'bold')).pack(pady=5, anchor="w")
    ttk.Button(frame_import_buttons, text="手动选择文件", 
               command=lambda: import_weekly_order_manual(global_data, status_var)).pack(fill=tk.X, pady=3)
    
    # 3. 每日订单导入
    ttk.Label(frame_import_buttons, text="每日订单导入", font=('微软雅黑', 12, 'bold')).pack(pady=5, anchor="w")
    ttk.Button(frame_import_buttons, text="自动下载并导入", 
               command=lambda: import_daily_order_manual(global_data, status_var)).pack(fill=tk.X, pady=3)
    
    # 3. 年预计量导入
    ttk.Label(frame_import_buttons, text="年预计量导入", font=("微软雅黑", 12, "bold")).pack(pady=5, anchor="w")
    ttk.Button(frame_import_buttons, text="手动选择文件", 
               command=lambda: import_yearly_estimate_manual(global_data, status_var)).pack(fill=tk.X, pady=3)
    
    # ========== 新增：已下单文件导入 ==========
    ttk.Label(frame_import_buttons, text="已下单文件导入", font=("微软雅黑", 12, "bold")).pack(pady=5, anchor="w")
    ttk.Button(frame_import_buttons, text="手动选择文件", 
               command=lambda: import_ordered_file_manual(global_data, status_var)).pack(fill=tk.X, pady=3)
    
    # 4. 当日成品库存导入
    ttk.Label(frame_import_buttons, text="当日成品库存导入", font=("微软雅黑", 12, "bold")).pack(pady=5, anchor="w")
    ttk.Button(frame_import_buttons, text="手动选择文件", 
               command=lambda: import_finished_stock_manual(global_data, status_var)).pack(fill=tk.X, pady=3)
    
    # 5. 当日毛坯库存导入
    ttk.Label(frame_import_buttons, text="当日毛坯库存导入", font=("微软雅黑", 12, "bold")).pack(pady=5, anchor="w")
    ttk.Button(frame_import_buttons, text="手动选择文件", 
               command=lambda: import_raw_stock_manual(global_data, status_var)).pack(fill=tk.X, pady=3)
    
    # ========== 新增：统一的导入所有之前文件按钮 ==========
    ttk.Separator(frame_import_buttons, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
    ttk.Button(frame_import_buttons, text="导入所有之前文件", 
               style="Accent.TButton",
               command=lambda: quick_import_all_last_files(global_data, status_var)).pack(fill=tk.X, pady=5)
    
    # ========== 中间填充空白（关键：让分析按钮固定在底部） ==========
    frame_spacer = ttk.Frame(frame_left)
    frame_spacer.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
    
    # ========== 底部分析按钮区域 ==========
    frame_analysis = ttk.Frame(frame_left)
    frame_analysis.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))
    
    # 分隔线
    ttk.Separator(frame_analysis, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=5)
    
    # 大字体分析按钮
    style = ttk.Style()
    style.configure("BigAccent.TButton", 
                    font=("微软雅黑", 14, "bold"),  # 放大字体
                    padding=10)  # 增加内边距
    
    # 先定义分析函数（避免提前引用analysis_widgets）
    def analysis_wrapper():
        # 从全局数据中获取analysis_widgets
        if "analysis_widgets" in global_data:
            run_analysis(global_data, path_config, status_var, global_data["analysis_widgets"])
        else:
            messagebox.showwarning("警告", "请先切换到「分析结果」标签页初始化界面！")
    
    ttk.Button(frame_analysis, text="开始采购分析", 
               style="BigAccent.TButton",
               command=analysis_wrapper).pack(fill=tk.X, padx=5, pady=5)
    
    # ========== 右侧数据概览区（新增已下单文件行） ==========
    frame_right = ttk.Frame(parent)
    frame_right.pack(side=tk.RIGHT, expand=1, fill="both", padx=10, pady=10)
    
    ttk.Label(frame_right, text="已导入数据概览", font=("微软雅黑", 12, "bold")).pack(pady=5)
    tree_overview = ttk.Treeview(frame_right, columns=("数据类型", "文件名", "数据量", "最后更新时间"), show="headings")
    tree_overview.heading("数据类型", text="数据类型")
    tree_overview.heading("文件名", text="文件名")
    tree_overview.heading("数据量", text="数据量")
    tree_overview.heading("最后更新时间", text="最后更新时间")
    tree_overview.column("数据类型", width=150)
    tree_overview.column("文件名", width=200)
    tree_overview.column("数据量", width=100)
    tree_overview.column("最后更新时间", width=200)
    tree_overview.pack(expand=1, fill="both")
    
    # 初始化概览表格（新增已下单文件行）
    tree_overview.insert("", "end", values=("毛坯基础信息", "未导入", "", "无"))
    tree_overview.insert("", "end", values=("每周订单", "未导入", "", "无"))
    tree_overview.insert("", "end", values=("每日订单", "未导入", "", "无"))  # 新增行
    tree_overview.insert("", "end", values=("年预计量", "未导入", "", "无"))
    tree_overview.insert("", "end", values=("已下单文件", "未导入", "", "无"))  # 新增行
    tree_overview.insert("", "end", values=("当日成品库存", "未导入", "", "无"))
    tree_overview.insert("", "end", values=("当日毛坯库存", "未导入", "", "无"))
    
    # 保存概览表格到全局数据
    global_data["tree_overview"] = tree_overview

def build_maintain_tab(parent, global_data):
    """构建订单追踪标签页"""
    ttk.Label(parent, text="订单追踪（已下单文件）", font=('微软雅黑', 12, 'bold')).pack(pady=10)
    
    # 定义表格列
    tree_maintain = ttk.Treeview(parent, columns=(
        "毛坯编码", "毛坯名称", "下单日期", "采购数量", "交货日期"
    ), show="headings")
    
    # 设置表头和列宽
    width_map = {
        "毛坯编码": 120,
        "毛坯名称": 150,
        "下单日期": 100,
        "采购数量": 80,
        "交货日期": 100
    }
    
    for col in tree_maintain["columns"]:
        tree_maintain.heading(col, text=col)
        tree_maintain.column(col, width=width_map.get(col, 100))
    
    tree_maintain.pack(expand=1, fill="both", padx=10, pady=10)
    global_data["tree_maintain"] = tree_maintain
    
    # 更新订单追踪表格
    update_order_tracking_table(global_data)


def update_order_tracking_table(global_data):
    """更新订单追踪表格"""
    tree = global_data.get("tree_maintain")
    if not tree:
        return
    
    # 清空表格
    for item in tree.get_children():
        tree.delete(item)
    
    # 获取已下单文件数据
    ordered_df = global_data.get("ordered_df")
    if ordered_df is None or ordered_df.empty:
        return
    
    # 直接使用原始数据，不进行过滤
    df = ordered_df.copy()
    
    # 过滤数据：只保留F列（实际交货时间）和G列（实际交货数量）都没有值的数据
    if "实际交货时间" in df.columns and "实际交货数量" in df.columns:
        # 过滤掉实际交货时间和实际交货数量都有值的行
        df = df[(df["实际交货时间"].isna() | (df["实际交货时间"].astype(str).str.strip() == '')) & 
                (df["实际交货数量"].isna() | (df["实际交货数量"].astype(str).str.strip() == ''))]
    
    # 按交货日期升序排序（最早的在前）
    if "交货日期" in df.columns:
        try:
            df["交货日期"] = pd.to_datetime(df["交货日期"], errors='coerce')
            df = df.sort_values(by="交货日期", ascending=True)
        except:
            pass
    
    # 为不同时间范围的订单创建标签
    tree.tag_configure('overdue', background='#FFCCCC', foreground='red')  # 延期：红色
    tree.tag_configure('urgent', background='#FFFFCC', foreground='black')  # 3天内：黄色
    tree.tag_configure('warning', background='#F5DEB3', foreground='black')  # 3-7天：棕色
    
    # 获取当前日期
    import datetime
    current_date = datetime.datetime.now().date()
    
    # 填充表格数据
    for _, row in df.iterrows():
        # 构建行数据
        row_data = []
        
        # 1. 毛坯编码 - B列物料列表
        raw_code = row.get("毛坯编码", "")
        row_data.append(str(raw_code))
        
        # 2. 毛坯名称 - C列物料名称
        raw_name = row.get("毛坯名称", "")
        row_data.append(str(raw_name))
        
        # 3. 下单日期 - A列采购日期
        order_date = row.get("下单日期", "")
        if pd.notna(order_date):
            if isinstance(order_date, pd.Timestamp):
                order_date = order_date.strftime('%Y-%m-%d')
            else:
                try:
                    date_obj = pd.to_datetime(order_date, errors='coerce')
                    if pd.notna(date_obj):
                        order_date = date_obj.strftime('%Y-%m-%d')
                    else:
                        order_date = str(order_date)
                except:
                    order_date = str(order_date)
        row_data.append(str(order_date))
        
        # 4. 采购数量 - D列采购数量
        quantity = row.get("采购数量", 0)
        if pd.notna(quantity):
            try:
                # 尝试将数量转换为数字
                if isinstance(quantity, str):
                    # 移除非数字字符
                    quantity = ''.join(filter(str.isdigit, quantity))
                quantity = str(int(float(quantity)))
            except:
                quantity = "0"
        else:
            quantity = "0"
        row_data.append(quantity)
        
        # 5. 交货日期 - E列预计交货日期
        delivery_date = row.get("交货日期", "")
        delivery_date_obj = None
        if pd.notna(delivery_date):
            if isinstance(delivery_date, pd.Timestamp):
                delivery_date_obj = delivery_date.date()
                delivery_date = delivery_date.strftime('%Y-%m-%d')
            else:
                try:
                    date_obj = pd.to_datetime(delivery_date, errors='coerce')
                    if pd.notna(date_obj):
                        delivery_date_obj = date_obj.date()
                        delivery_date = date_obj.strftime('%Y-%m-%d')
                    else:
                        delivery_date = str(delivery_date)
                except:
                    delivery_date = str(delivery_date)
        row_data.append(str(delivery_date))
        
        # 判断时间范围并标记
        tags = []
        if delivery_date_obj:
            # 计算交货日期与当前日期的天数差
            days_diff = (delivery_date_obj - current_date).days
            
            if days_diff < 0:
                # 延期：红色
                tags.append('overdue')
            elif days_diff <= 3:
                # 3天内：黄色
                tags.append('urgent')
            elif days_diff < 7:
                # 3-7天：棕色
                tags.append('warning')
        
        # 插入行
        tree.insert("", "end", values=row_data, tags=tags)