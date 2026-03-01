# ui_analysis.py - 采购分析结果展示模块（成品维度版，修复ttk Treeview字体报错）
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

def build_analysis_tab(parent, global_data, path_config):
    """
    构建分析结果标签页（成品维度版）
    表格字段：成品编码、成品名称、毛坯编码、毛坯名称、订单数量、年预计量、成品库存、毛坯库存、已下单数量、库存状态
    修复：ttk.Treeview表头字体不支持-font参数的问题
    """
    widgets = {}
    
    # ========== 修复：配置ttk Treeview样式（解决表头字体问题） ==========
    style = ttk.Style()
    # 定义Treeview表头样式
    style.configure("Custom.Treeview.Heading", font=("微软雅黑", 9, "bold"))
    # 定义Treeview内容样式（使用微软雅黑字体，优化行高）
    style.configure("Custom.Treeview", font=("微软雅黑", 9), rowheight=18)
    # 性能优化：配置选中样式（蓝色背景，黑色文字）
    # 简化样式映射，减少渲染开销
    style.map("Custom.Treeview", 
              background=[('selected', '#0078D7')],
              foreground=[('selected', 'black')])
    
    # ========== 1. 预警卡片区域 ==========
    frame_cards = ttk.Frame(parent)
    frame_cards.pack(fill=tk.X, padx=5, pady=5)
    
    # 需下单卡片（红色）
    card_need_purchase = ttk.Label(
        frame_cards, 
        text="需下单：0件", 
        font=("微软雅黑", 12, "bold"),
        background="#ff4444", 
        foreground="white"
    )
    card_need_purchase.pack(side=tk.LEFT, expand=1, fill=tk.X, padx=2, pady=2)
    widgets["card_need_purchase"] = card_need_purchase
    
    # 需警惕卡片（黄色）
    card_need_warning = ttk.Label(
        frame_cards, 
        text="需警惕：0件", 
        font=("微软雅黑", 12, "bold"),
        background="#ffff00", 
        foreground="black"
    )
    card_need_warning.pack(side=tk.LEFT, expand=1, fill=tk.X, padx=2, pady=2)
    widgets["card_need_warning"] = card_need_warning
    
    # 正常卡片（绿色）
    card_normal = ttk.Label(
        frame_cards, 
        text="正常：0件", 
        font=("微软雅黑", 12, "bold"),
        background="#00cc00", 
        foreground="white"
    )
    card_normal.pack(side=tk.LEFT, expand=1, fill=tk.X, padx=2, pady=2)
    widgets["card_normal"] = card_normal
    
    # ========== 2. 结果表格区域（成品维度） ==========
    frame_table = ttk.Frame(parent)
    frame_table.pack(expand=1, fill="both", padx=5, pady=5)
    
    # 表格标题
    ttk.Label(frame_table, text="采购分析结果详情（成品维度）", font=("微软雅黑", 10, "bold")).pack(pady=2)
    
    # 定义表格列（去除成品名称、总订单和库存状态）
    table_columns = (
        "成品编码", "毛坯编码", "毛坯名称",
        "年预计量", "精准季预计量", "今日订单", "成品库存", "毛坯库存", "已下单数量"
    )
    
    # 创建表格（使用自定义样式）
    tree_analysis = ttk.Treeview(frame_table, columns=table_columns, show="headings", style="Custom.Treeview")
    
    # 性能优化：禁用Treeview的一些默认行为
    tree_analysis.configure(
        selectmode="browse",  # 只允许选择一行
        takefocus=True        # 允许获取焦点
    )
    
    # 设置表头和列宽（优化列宽以适应一页显示）
    width_map = {
        "成品编码": 120,  # 调整为120px
        "毛坯编码": 160,  # 调整为160px
        "毛坯名称": 120,
        "年预计量": 70,
        "精准季预计量": 90,
        "今日订单": 70,
        "成品库存": 70,
        "毛坯库存": 70,
        "已下单数量": 80
    }
    for col in table_columns:
        tree_analysis.heading(col, text=col)
        # 性能优化：禁用列的最小宽度限制
        # 对于毛坯名称列，允许它自动拉伸以填充剩余空间
        if col == "毛坯名称":
            tree_analysis.column(col, width=width_map[col], minwidth=0, stretch=True)
        else:
            tree_analysis.column(col, width=width_map[col], minwidth=0, stretch=False)
    
    # 性能优化：只添加垂直滚动条，禁用横向滚动条
    vsb = ttk.Scrollbar(frame_table, orient="vertical", command=tree_analysis.yview)
    tree_analysis.configure(yscrollcommand=vsb.set, xscrollcommand="")
    
    # 布局（使用pack）
    vsb.pack(side=tk.RIGHT, fill=tk.Y)
    tree_analysis.pack(side=tk.LEFT, expand=1, fill="both")
    
    # ========== 4. 添加右键菜单功能（复制指定列） ==========
    # 存储右键点击的列索引
    clicked_column_index = [0]  # 使用列表存储，以便在回调中修改
    
    # 创建右键菜单
    right_click_menu = tk.Menu(tree_analysis, tearoff=0)
    
    # 复制功能实现
    def copy_selected_column_content(tree):
        selected_item = tree.selection()[0] if tree.selection() else None
        if selected_item:
            # 获取选中行的所有列值
            values = tree.item(selected_item, "values")
            # 获取右键点击的列索引
            col_index = clicked_column_index[0]
            if 0 <= col_index < len(values):
                # 获取对应列的值
                content = values[col_index]
                # 复制到剪贴板
                root = tree.winfo_toplevel()
                root.clipboard_clear()
                root.clipboard_append(str(content))
                root.update()  # 确保内容被复制
    
    # 右键菜单回调函数
    def show_right_click_menu(event):
        # 确保有选中的项目
        if tree_analysis.selection():
            try:
                # 获取右键点击的位置对应的列索引
                region = tree_analysis.identify("region", event.x, event.y)
                if region == "cell":
                    # 计算列索引
                    x = event.x
                    col_widths = [tree_analysis.column(col, "width") for col in table_columns]
                    total_width = 0
                    col_index = 0
                    for i, width in enumerate(col_widths):
                        total_width += width
                        if x < total_width:
                            col_index = i
                            break
                    
                    # 更新存储的列索引
                    clicked_column_index[0] = col_index
                    
                    # 获取列名
                    col_name = table_columns[col_index]
                    
                    # 更新右键菜单的标签，显示当前列名
                    right_click_menu.delete(0, tk.END)  # 清空菜单
                    right_click_menu.add_command(
                        label=f"复制 {col_name} 值", 
                        command=lambda: copy_selected_column_content(tree_analysis)
                    )
                    
                    # 显示菜单
                    right_click_menu.tk_popup(event.x_root, event.y_root)
            finally:
                # 释放菜单
                right_click_menu.grab_release()
    
    # 绑定右键点击事件
    tree_analysis.bind("<Button-3>", show_right_click_menu)
    
    # 绑定双击事件
    def on_double_click(event):
        # 获取双击的行
        selected_item = tree_analysis.selection()
        if not selected_item:
            return
        
        # 获取选中行的数据
        values = tree_analysis.item(selected_item[0], "values")
        
        # 找到成品编码和毛坯编码的索引
        # 成品编码在第1列（索引0），毛坯编码在第2列（索引1）
        finished_code_index = 0
        blank_code_index = 1
        
        if finished_code_index < len(values) and blank_code_index < len(values):
            finished_code = values[finished_code_index]
            blank_code = values[blank_code_index]
            if blank_code and blank_code != "-":
                # 打开新窗口显示该成品编码和毛坯编码的订单记录
                show_blank_code_orders(blank_code, global_data, finished_code)
    
    tree_analysis.bind("<Double-Button-1>", on_double_click)
    
    widgets["tree_analysis"] = tree_analysis
    widgets["table_columns"] = table_columns
    
    # ========== 3. 操作按钮区域 ==========
    frame_buttons = ttk.Frame(parent)
    frame_buttons.pack(fill=tk.X, padx=5, pady=5)
    
    # 筛选急需采购按钮
    def filter_urgent():
        update_analysis_table(widgets, global_data, filter_type="urgent")
    
    # 重置显示全部按钮
    def reset_filter():
        update_analysis_table(widgets, global_data, filter_type="all")
    
    # 导出Excel按钮
    def export_excel():
        try:
            analysis_df = global_data.get("analysis_show_df")
            if analysis_df is None or analysis_df.empty:
                messagebox.showwarning("提示", "暂无分析数据可导出！")
                return
            
            # 仅导出表格中显示的字段（重命名为表格列名）
            export_columns = [
                "成品物料编码", "成品物料名称", "毛坯物料编码", "毛坯物料名称",
                "年预计量", "精准季预计量", "订单数量", "成品库存数量", "毛坯库存数量", "已下单数量", "库存健康状态"
            ]
            export_df = analysis_df[export_columns].copy()
            
            # 重命名列（和表格一致）
            export_df.rename(columns={
                "成品物料编码": "成品编码",
                "成品物料名称": "成品名称",
                "毛坯物料编码": "毛坯编码",
                "毛坯物料名称": "毛坯名称",
                "成品库存数量": "成品库存",
                "毛坯库存数量": "毛坯库存",
                "库存健康状态": "库存状态"
            }, inplace=True)
            
            # 选择保存路径
            from tkinter import filedialog
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
                title="导出采购分析结果（成品维度）"
            )
            if save_path:
                export_df.to_excel(save_path, index=False)
                messagebox.showinfo("成功", f"分析结果已导出到：\n{save_path}")
        except Exception as e:
            messagebox.showerror("错误", f"导出失败：{str(e)}")
    
    # 按钮布局
    ttk.Button(frame_buttons, text="筛选急需采购", command=filter_urgent, width=10).pack(side=tk.LEFT, padx=2)
    ttk.Button(frame_buttons, text="显示全部", command=reset_filter, width=8).pack(side=tk.LEFT, padx=2)
    ttk.Button(frame_buttons, text="导出Excel", command=export_excel, width=8).pack(side=tk.RIGHT, padx=2)
    
    return widgets

def update_analysis_table(widgets, global_data, filter_type="all"):
    """
    更新分析结果表格数据（成品维度版）
    filter_type: all-全部, urgent-仅需下单
    性能优化：使用values批量插入，禁用重绘，优化pandas操作
    """
    tree_analysis = widgets["tree_analysis"]
    
    # 设置列对齐方式为居中
    for col in tree_analysis["columns"]:
        tree_analysis.column(col, anchor="center")
    
    # 性能优化：禁用自动更新和重绘
    tree_analysis.configure(style="Custom.Treeview")
    tree_analysis.state(["disabled"])
    
    try:
        # 清空表格（更高效的方式）
        items = tree_analysis.get_children()
        if items:
            tree_analysis.delete(*items)
        
        # 获取分析数据
        analysis_df = global_data.get("analysis_show_df")
        if analysis_df is None or analysis_df.empty:
            tree_analysis.insert("", "end", values=["暂无数据"] * len(widgets["table_columns"]))
            return
        
        # 筛选数据
        if filter_type == "urgent":
            display_df = analysis_df[analysis_df["建议下单数量"] > 0].copy()
        else:
            display_df = analysis_df.copy()
        
        # 去重逻辑：基于所有列去重
        # 这样可以避免完全相同的记录重复出现，同时保留不同的记录
        # 即使物料编码为空，也能够显示所有的记录
        if not display_df.empty:
            display_df = display_df.drop_duplicates(
                keep="first"
            )
        
        # 字段映射
        column_mapping = {
            "成品编码": "成品物料编码",
            "毛坯编码": "毛坯物料编码",
            "毛坯名称": "毛坯物料名称",
            "年预计量": "年预计量",
            "精准季预计量": "精准季预计量",
            "今日订单": "今日订单数量",
            "成品库存": "成品库存数量",
            "毛坯库存": "毛坯库存数量",
            "已下单数量": "已下单数量"
        }
        
        # 性能优化：使用向量化操作替代iterrows
        numeric_fields = ["今日订单数量", "年预计量", "精准季预计量", "成品库存数量", "毛坯库存数量", "已下单数量"]
        for field in numeric_fields:
            if field in display_df.columns:
                display_df[field] = pd.to_numeric(display_df[field], errors="coerce").fillna(0).astype(int)
        
        if "颜色状态" not in display_df.columns:
            display_df["颜色状态"] = "normal"
        
        # 检查是否为空DataFrame
        if display_df.empty:
            tree_analysis.state(["!disabled"])
            return
        
        # 性能优化：使用向量化操作构建所有数据
        table_cols = widgets["table_columns"]
        
        # 更高效的数据处理：使用列表推导式替代apply
        all_values = []
        all_tags = []
        
        # 预提取所需列的数据
        required_cols = [column_mapping[col] for col in table_cols]
        color_states = display_df["颜色状态"].tolist()
        
        # 记录上一个数值字段的值和上一个毛坯编码，用于合并单元格
        last_blank_code = None
        last_yearly_estimate = float('-inf')  # 使用负无穷大作为初始值
        last_quarterly_estimate = float('-inf')  # 使用负无穷大作为初始值
        last_today_order = float('-inf')  # 使用负无穷大作为初始值
        last_finished_stock = float('-inf')  # 使用负无穷大作为初始值
        last_blank_stock = float('-inf')  # 使用负无穷大作为初始值
        last_ordered = float('-inf')  # 使用负无穷大作为初始值
        
        # 使用iloc进行更高效的行访问
        for idx in range(len(display_df)):
            row = display_df.iloc[idx]
            row_values = []
            
            # 获取当前行的相关字段的值
            current_blank_code = row.get("毛坯物料编码", "")
            current_yearly_estimate = row.get("年预计量", 0)
            current_quarterly_estimate = row.get("精准季预计量", 0)
            current_today_order = row.get("今日订单数量", 0)
            current_finished_stock = row.get("成品库存数量", 0)
            current_blank_stock = row.get("毛坯库存数量", 0)
            current_ordered = row.get("已下单数量", 0)
            
            # 检查当前毛坯编码是否与上一个相同
            if current_blank_code != last_blank_code:
                # 不同毛坯编码，重置所有last值
                last_blank_code = current_blank_code
                last_yearly_estimate = float('-inf')
                last_quarterly_estimate = float('-inf')
                last_today_order = float('-inf')
                last_finished_stock = float('-inf')
                last_blank_stock = float('-inf')
                last_ordered = float('-inf')
            
            # 遍历所有列
            for col in required_cols:
                # 检查是否是需要合并的列
                if col == "年预计量":
                    # 如果当前年预计量与上一个相同，显示为空，否则显示当前年预计量
                    if current_yearly_estimate == last_yearly_estimate:
                        val = ""
                    else:
                        val = str(current_yearly_estimate)
                        last_yearly_estimate = current_yearly_estimate
                elif col == "精准季预计量":
                    # 如果当前精准季预计量与上一个相同，显示为空，否则显示当前精准季预计量
                    if current_quarterly_estimate == last_quarterly_estimate:
                        val = ""
                    else:
                        val = str(current_quarterly_estimate)
                        last_quarterly_estimate = current_quarterly_estimate
                elif col == "今日订单数量":
                    # 如果当前今日订单数量与上一个相同，显示为空，否则显示当前今日订单数量
                    if current_today_order == last_today_order:
                        val = ""
                    else:
                        val = str(current_today_order)
                        last_today_order = current_today_order
                elif col == "成品库存数量":
                    # 如果当前成品库存数量与上一个相同，显示为空，否则显示当前成品库存数量
                    if current_finished_stock == last_finished_stock:
                        val = ""
                    else:
                        val = str(current_finished_stock)
                        last_finished_stock = current_finished_stock
                elif col == "毛坯库存数量":
                    # 如果当前毛坯库存数量与上一个相同，显示为空，否则显示当前毛坯库存数量
                    if current_blank_stock == last_blank_stock:
                        val = ""
                    else:
                        val = str(current_blank_stock)
                        last_blank_stock = current_blank_stock
                elif col == "已下单数量":
                    # 如果当前已下单数量与上一个相同，显示为空，否则显示当前已下单数量
                    if current_ordered == last_ordered:
                        val = ""
                    else:
                        val = str(current_ordered)
                        last_ordered = current_ordered
                else:
                    # 其他列按原来的逻辑处理
                    val = row.get(col, "")
                    if pd.isna(val):
                        val = ""
                    elif val == "":
                        val = ""
                    elif isinstance(val, (int, float)):
                        val = str(int(val))
                    else:
                        val = str(val)
                row_values.append(val)
            all_values.append(row_values)
            all_tags.append(color_states[idx] if color_states[idx] in ("red", "yellow", "brown") else "")
        
        # 性能优化：批量插入所有行（使用更高效的方式）
        # 预配置标签（只做一次）
        if not hasattr(tree_analysis, "_tags_configured"):
            tree_analysis.tag_configure("red", background="#FFCCCC")
            tree_analysis.tag_configure("yellow", background="#FFFFCC")
            tree_analysis.tag_configure("brown", background="#D2691E")
            tree_analysis._tags_configured = True
        
        # 批量插入数据
        for values, tag in zip(all_values, all_tags):
            tags = (tag,) if tag else ()
            tree_analysis.insert("", "end", values=values, tags=tags)
            
    finally:
        # 性能优化：重新启用更新
        tree_analysis.state(["!disabled"])
        # 强制刷新一次，确保所有数据显示
        tree_analysis.update_idletasks()

def update_analysis_cards(widgets, global_data):
    """更新预警卡片数据（成品维度）"""
    analysis_df = global_data.get("analysis_show_df")
    if analysis_df is None or analysis_df.empty:
        widgets["card_need_purchase"].config(text="需下单：0件")
        widgets["card_need_warning"].config(text="需警惕：0件")
        widgets["card_normal"].config(text="正常：0件")
        return
    
    # 确保颜色状态列存在
    if "颜色状态" not in analysis_df.columns:
        analysis_df["颜色状态"] = "normal"
    
    # 统计各状态数量（基于颜色状态）
    urgent_count = len(analysis_df[analysis_df["颜色状态"] == "red"])
    warning_count = len(analysis_df[analysis_df["颜色状态"] == "yellow"])
    normal_count = len(analysis_df[analysis_df["颜色状态"] == "normal"])
    
    widgets["card_need_purchase"].config(text=f"需下单：{urgent_count}件")
    widgets["card_need_warning"].config(text=f"需警惕：{warning_count}件")
    widgets["card_normal"].config(text=f"正常：{normal_count}件")

def show_blank_code_orders(blank_code, global_data, finished_code=None):
    """
    显示指定毛坯编码的订单记录和预计采购分析（新窗口）
    参数：
        blank_code: 毛坯编码
        global_data: 全局数据
        finished_code: 成品编码（可选，用于精确匹配记录）
    """
    # 创建新窗口
    new_window = tk.Toplevel()
    new_window.title(f"毛坯编码【{blank_code}】详情")
    new_window.geometry("1000x600")
    new_window.resizable(True, True)
    
    # 居中显示
    new_window.update_idletasks()
    width = new_window.winfo_width()
    height = new_window.winfo_height()
    screen_width = new_window.winfo_screenwidth()
    screen_height = new_window.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    new_window.geometry(f'{width}x{height}+{x}+{y}')
    
    # 创建标签页控件
    notebook = ttk.Notebook(new_window)
    notebook.pack(expand=1, fill="both")
    
    # ========== 1. 预采购分析标签页 ==========
    tab_analysis = ttk.Frame(notebook)
    notebook.add(tab_analysis, text="预采购分析")
    
    # 创建主容器，使用网格布局
    main_frame = ttk.Frame(tab_analysis)
    main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    # 获取当前日期和周数
    import datetime
    from analysis_engine import calculate_week_number
    
    current_date = datetime.datetime.now()
    date_str = current_date.strftime("%Y年%m月%d日")
    current_week_str = calculate_week_number(current_date.strftime("%Y-%m-%d"))
    # 格式化为w2605格式（与年预计量文件匹配）
    current_week_format = f"w{current_week_str[2:4]}{current_week_str[5:]}"
    
    # 从分析数据中获取该毛坯编码的相关信息
    analysis_df = global_data.get("analysis_show_df")
    analysis_record = None
    if analysis_df is not None and not analysis_df.empty:
        for idx, row in analysis_df.iterrows():
            if '毛坯物料编码' in row and str(row['毛坯物料编码']).strip() == blank_code.strip():
                if finished_code is not None:
                    if '成品物料编码' in row and str(row['成品物料编码']).strip() == finished_code.strip():
                        analysis_record = row
                        break
                else:
                    analysis_record = row
                    break
    
    # 1. 当前日期和周数区域
    date_frame = ttk.LabelFrame(main_frame, text="当前信息", padding=(10, 5))
    date_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=5)
    
    # 创建内部横向排列容器
    date_inner_frame = ttk.Frame(date_frame)
    date_inner_frame.pack(fill=tk.BOTH, expand=True)
    
    # 当前日期
    date_card = ttk.Frame(date_inner_frame)
    date_card.pack(side=tk.LEFT, padx=20, pady=5)
    date_label = ttk.Label(date_card, text="当前日期", font=("微软雅黑", 11, "bold"))
    date_label.pack()
    date_value = ttk.Label(date_card, text=date_str, font=("微软雅黑", 10))
    date_value.pack()
    
    # 当前周
    week_card = ttk.Frame(date_inner_frame)
    week_card.pack(side=tk.LEFT, padx=20, pady=5)
    week_label = ttk.Label(week_card, text="当前周", font=("微软雅黑", 11, "bold"))
    week_label.pack()
    week_value = ttk.Label(week_card, text=current_week_format, font=("微软雅黑", 10))
    week_value.pack()
    
    # 2. 基本信息区域
    basic_frame = ttk.LabelFrame(main_frame, text="基本信息", padding=(10, 5))
    basic_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=5)
    
    # 创建内部横向排列容器
    basic_inner_frame = ttk.Frame(basic_frame)
    basic_inner_frame.pack(fill=tk.BOTH, expand=True)
    
    if analysis_record is not None:
        # 成品物料编码
        finished_code_card = ttk.Frame(basic_inner_frame)
        finished_code_card.pack(side=tk.LEFT, padx=15, pady=5)
        finished_code_label = ttk.Label(finished_code_card, text="成品编码", font=("微软雅黑", 11, "bold"))
        finished_code_label.pack()
        finished_code_value = ttk.Label(finished_code_card, text=str(analysis_record.get("成品物料编码", "-")).strip(), font=("微软雅黑", 10))
        finished_code_value.pack()
        
        # 毛坯物料编码
        blank_code_card = ttk.Frame(basic_inner_frame)
        blank_code_card.pack(side=tk.LEFT, padx=15, pady=5)
        blank_code_label = ttk.Label(blank_code_card, text="毛坯编码", font=("微软雅黑", 11, "bold"))
        blank_code_label.pack()
        blank_code_value = ttk.Label(blank_code_card, text=str(analysis_record.get("毛坯物料编码", "-")).strip(), font=("微软雅黑", 10))
        blank_code_value.pack()
        
        # 毛坯物料名称
        blank_name_card = ttk.Frame(basic_inner_frame)
        blank_name_card.pack(side=tk.LEFT, padx=15, pady=5)
        blank_name_label = ttk.Label(blank_name_card, text="毛坯名称", font=("微软雅黑", 11, "bold"))
        blank_name_label.pack()
        blank_name_value = ttk.Label(blank_name_card, text=str(analysis_record.get("毛坯物料名称", "-")).strip(), font=("微软雅黑", 10))
        blank_name_value.pack()
    
    # 3. 预计量与库存区域
    stock_frame = ttk.LabelFrame(main_frame, text="预计量与库存", padding=(10, 5))
    stock_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=5)
    
    # 创建内部横向排列容器
    stock_inner_frame = ttk.Frame(stock_frame)
    stock_inner_frame.pack(fill=tk.BOTH, expand=True)
    
    if analysis_record is not None:
        # 年预计量
        yearly_card = ttk.Frame(stock_inner_frame)
        yearly_card.pack(side=tk.LEFT, padx=15, pady=5)
        yearly_label = ttk.Label(yearly_card, text="年预计量", font=("微软雅黑", 11, "bold"))
        yearly_label.pack()
        yearly_value = ttk.Label(yearly_card, text=str(analysis_record.get("年预计量", "-")).strip(), font=("微软雅黑", 10))
        yearly_value.pack()
        
        # 精准季预计量
        quarter_card = ttk.Frame(stock_inner_frame)
        quarter_card.pack(side=tk.LEFT, padx=15, pady=5)
        quarter_label = ttk.Label(quarter_card, text="精准季预计量", font=("微软雅黑", 11, "bold"))
        quarter_label.pack()
        quarter_value = ttk.Label(quarter_card, text=str(analysis_record.get("精准季预计量", "-")).strip(), font=("微软雅黑", 10))
        quarter_value.pack()
        
        # 成品库存数量
        finished_stock_card = ttk.Frame(stock_inner_frame)
        finished_stock_card.pack(side=tk.LEFT, padx=15, pady=5)
        finished_stock_label = ttk.Label(finished_stock_card, text="成品库存", font=("微软雅黑", 11, "bold"))
        finished_stock_label.pack()
        finished_stock_value = ttk.Label(finished_stock_card, text=str(analysis_record.get("成品库存数量", "-")).strip(), font=("微软雅黑", 10))
        finished_stock_value.pack()
        
        # 毛坯库存数量
        blank_stock_card = ttk.Frame(stock_inner_frame)
        blank_stock_card.pack(side=tk.LEFT, padx=15, pady=5)
        blank_stock_label = ttk.Label(blank_stock_card, text="毛坯库存", font=("微软雅黑", 11, "bold"))
        blank_stock_label.pack()
        blank_stock_value = ttk.Label(blank_stock_card, text=str(analysis_record.get("毛坯库存数量", "-")).strip(), font=("微软雅黑", 10))
        blank_stock_value.pack()
    
    # 4. 安全库存设置区域
    safety_frame = ttk.LabelFrame(main_frame, text="安全库存设置", padding=(10, 5))
    safety_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=5)
    
    # 创建内部横向排列容器
    safety_inner_frame = ttk.Frame(safety_frame)
    safety_inner_frame.pack(fill=tk.BOTH, expand=True)
    
    # 获取成品编码（优先使用函数参数，其次使用分析记录中的值）
    if finished_code is None or finished_code == "":
        finished_code = "-"
    if analysis_record is not None:
        finished_code = str(analysis_record.get("成品物料编码", "-")).strip()
    
    # 查询基础信息表获取建议安全周数
    raw_material_df = global_data.get("raw_material_df")
    safety_weeks = 4  # 默认4周
    if raw_material_df is not None and not raw_material_df.empty:
        for idx, row in raw_material_df.iterrows():
            if '成品物料编码' in row and str(row['成品物料编码']).strip() == finished_code:
                # 获取O列（索引14）的建议安全库存周
                if len(row) > 14:
                    safety_weeks_val = row.iloc[14]
                    try:
                        safety_weeks = int(safety_weeks_val)
                        if safety_weeks <= 0:
                            safety_weeks = 4
                    except:
                        safety_weeks = 4
                break
    
    # 建议安全库存周
    safety_card = ttk.Frame(safety_inner_frame)
    safety_card.pack(side=tk.LEFT, padx=15, pady=5)
    safety_label = ttk.Label(safety_card, text="建议安全库存周", font=("微软雅黑", 11, "bold"))
    safety_label.pack()
    safety_value = ttk.Label(safety_card, text=str(safety_weeks) + "周", font=("微软雅黑", 10))
    safety_value.pack()
    
    # 5. 预采购量周数数据区域
    purchase_frame = ttk.LabelFrame(main_frame, text="预采购量周数数据", padding=(10, 5))
    purchase_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=5)
    
    # 创建一个内部框架用于横向排列
    purchase_inner_frame = ttk.Frame(purchase_frame)
    purchase_inner_frame.pack(fill=tk.BOTH, expand=True)
    
    # 获取年预计量文件，查询周数预采购量
    yearly_estimate_df = global_data.get("yearly_estimate_df")
    weekly_data = []
    
    if yearly_estimate_df is not None and not yearly_estimate_df.empty:
        # 首先尝试使用毛坯编码查找
        for idx, row in yearly_estimate_df.iterrows():
            if len(row) > 1:
                row_blank_code = str(row.iloc[1]).strip()
                if row_blank_code == blank_code:
                    # 尝试处理更多列范围的周数数据
                    for col_idx in range(len(row)):
                        col_name = yearly_estimate_df.columns[col_idx]
                        if isinstance(col_name, str):
                            # 检查列名是否包含周数信息（支持w2605格式）
                            col_name_str = str(col_name).strip()
                            if col_name_str.startswith('w') or col_name_str.startswith('W'):
                                week_value = row.iloc[col_idx]
                                try:
                                    week_value = int(week_value)
                                except:
                                    week_value = 0
                                weekly_data.append([col_name_str, week_value])
                    break
        
        # 如果使用毛坯编码找不到，尝试使用成品编码查找
        if not weekly_data and finished_code != "-":
            for idx, row in yearly_estimate_df.iterrows():
                if len(row) > 2:
                    # 查找C列（索引2）的物料代码
                    row_finished_code = str(row.iloc[2]).strip()
                    if row_finished_code == finished_code:
                        # 尝试处理更多列范围的周数数据
                        for col_idx in range(len(row)):
                            col_name = yearly_estimate_df.columns[col_idx]
                            if isinstance(col_name, str):
                                # 检查列名是否包含周数信息（支持w2605格式）
                                col_name_str = str(col_name).strip()
                                if col_name_str.startswith('w') or col_name_str.startswith('W'):
                                    week_value = row.iloc[col_idx]
                                    try:
                                        week_value = int(week_value)
                                    except:
                                        week_value = 0
                                    weekly_data.append([col_name_str, week_value])
                        break
    
    # 显示预采购量数据
    if weekly_data:
        # 找到当前周的索引
        current_week_idx = -1
        for i, (week_code, _) in enumerate(weekly_data):
            if isinstance(week_code, str) and isinstance(current_week_format, str):
                if week_code.upper() == current_week_format.upper():
                    current_week_idx = i
                    break
        
        # 展示当前周及之后的安全库存周数量的数据
        if current_week_idx != -1:
            # 显示当前周
            current_week_data = weekly_data[current_week_idx]
            week_num = current_week_data[0][-2:]  # 获取周数数字，如"05"
            
            # 创建当前周数据卡片
            current_week_card = ttk.Frame(purchase_inner_frame)
            current_week_card.pack(side=tk.LEFT, padx=10, pady=5)
            
            current_week_label = ttk.Label(current_week_card, text=f"{week_num}周预采购量", font=("微软雅黑", 11, "bold"))
            current_week_label.pack()
            current_week_value = ttk.Label(current_week_card, text=str(current_week_data[1]), font=("微软雅黑", 10))
            current_week_value.pack()
            
            # 显示之后的安全库存周（从下一周开始，不包含本周）
            for i in range(current_week_idx + 1, min(current_week_idx + 1 + safety_weeks, len(weekly_data))):
                week_code, week_value = weekly_data[i]
                week_num = week_code[-2:]  # 获取周数数字
                
                # 创建周数据卡片
                week_card = ttk.Frame(purchase_inner_frame)
                week_card.pack(side=tk.LEFT, padx=10, pady=5)
                
                week_label = ttk.Label(week_card, text=f"{week_num}周预采购量", font=("微软雅黑", 11, "bold"))
                week_label.pack()
                week_value_label = ttk.Label(week_card, text=str(week_value), font=("微软雅黑", 10))
                week_value_label.pack()
        else:
            # 如果找不到当前周，至少显示第一条数据
            if weekly_data:
                first_week_data = weekly_data[0]
                week_num = first_week_data[0][-2:]  # 获取周数数字
                
                # 创建周数据卡片
                first_week_card = ttk.Frame(purchase_inner_frame)
                first_week_card.pack(side=tk.LEFT, padx=10, pady=5)
                
                first_week_label = ttk.Label(first_week_card, text=f"{week_num}周预采购量", font=("微软雅黑", 11, "bold"))
                first_week_label.pack()
                first_week_value = ttk.Label(first_week_card, text=str(first_week_data[1]), font=("微软雅黑", 10))
                first_week_value.pack()
    else:
        # 无预采购量数据提示
        no_data_label = ttk.Label(purchase_inner_frame, text="暂无预采购量数据", font=("微软雅黑", 10))
        no_data_label.pack(padx=15, pady=10, anchor="w")
    
    # ========== 2. 订单记录标签页 ==========
    tab_orders = ttk.Frame(notebook)
    notebook.add(tab_orders, text="订单记录")
    
    # 检查已下单文件是否存在
    ordered_df = global_data.get("ordered_df")
    if ordered_df is None or ordered_df.empty:
        ttk.Label(tab_orders, text="未找到已下单文件数据！", font=("微软雅黑", 10)).pack(pady=20)
    else:
        # 根据毛坯编码过滤数据
        filtered_df = ordered_df[ordered_df['毛坯编码'].str.strip() == blank_code.strip()].copy()
        
        if filtered_df.empty:
            ttk.Label(tab_orders, text=f"未找到毛坯编码【{blank_code}】的订单记录！", font=("微软雅黑", 10)).pack(pady=20)
        else:
            # 创建表格列（去除采购日期和成品编码列）
            exclude_columns = ['采购日期', '成品编码']
            table_columns = [col for col in ordered_df.columns if col not in exclude_columns]
            
            # 配置表格样式
            style = ttk.Style()
            style.configure("Order.Treeview", font=("微软雅黑", 9))
            style.configure("Order.Treeview.Heading", font=("微软雅黑", 9, "bold"))
            
            tree_orders = ttk.Treeview(tab_orders, columns=table_columns, show="headings", style="Order.Treeview")
            
            # 设置表头和列宽（所有列居中对齐）
            for col in table_columns:
                tree_orders.heading(col, text=col)
                tree_orders.column(col, width=100, minwidth=50, anchor='center')
            
            # 添加滚动条
            vsb_orders = ttk.Scrollbar(tab_orders, orient="vertical", command=tree_orders.yview)
            tree_orders.configure(yscrollcommand=vsb_orders.set)
            vsb_orders.pack(side=tk.RIGHT, fill=tk.Y)
            tree_orders.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            # 填充数据
            for _, row in filtered_df.iterrows():
                values = []
                for col in table_columns:
                    val = row.get(col, "")
                    if pd.isna(val):
                        val = ""
                    elif isinstance(val, (int, float)):
                        val = f"{val}"
                    else:
                        # 处理日期列（下单日期、交货日期、实际交货时间），格式化为年月日
                        if col in ['下单日期', '交货日期', '实际交货时间']:
                            try:
                                # 尝试转换为日期并格式化
                                date_val = pd.to_datetime(val, errors='coerce')
                                if pd.notna(date_val):
                                    val = date_val.strftime('%Y-%m-%d')
                                else:
                                    val = ""
                            except:
                                val = ""
                        else:
                            val = str(val)
                    values.append(val)
                tree_orders.insert("", "end", values=values)
