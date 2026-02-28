# ui_analysis_tksheet.py - 使用tksheet虚拟化表格的高性能分析结果展示模块
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

try:
    import tksheet
    TKSHEET_AVAILABLE = True
except ImportError:
    TKSHEET_AVAILABLE = False

def build_analysis_tab(parent, global_data, path_config):
    """
    构建分析结果标签页（使用tksheet虚拟化表格）
    """
    widgets = {}
    
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
    
    # ========== 2. 结果表格区域 ==========
    frame_table = ttk.Frame(parent)
    frame_table.pack(expand=1, fill="both", padx=5, pady=5)
    
    # 表格标题
    ttk.Label(frame_table, text="采购分析结果详情（成品维度）", font=("微软雅黑", 10, "bold")).pack(pady=2)
    
    # 定义表格列（去除成品名称）
    table_columns = (
        "成品编码", "毛坯编码", "毛坯名称",
        "年预计量", "精准季预计量", "总订单", "今日订单", "成品库存", "毛坯库存", "已下单数量", "库存状态"
    )
    
    # 创建tksheet
    if TKSHEET_AVAILABLE:
        # 创建Sheet实例（简化配置）
        sheet = tksheet.Sheet(
            frame_table,
            data=[[""] * len(table_columns)],
            headers=table_columns,
            show_row_index=False,
            show_column_index=False,
            enable_edit_events=False,
            enable_edit_cell=False,
            row_height=18,  # 减小行高以提升性能
            default_row_height=18,
            show_x_scrollbar=False,  # 禁用横向滚动条
            show_y_scrollbar=True,
            page_up_down_select_row=True,
            display_selected_cells_with_highlights=False,  # 禁用高亮以提升性能
            selected_rows_to_end_of_window=False,  # 禁用以提升性能
            arrow_key_down_selects_row=True,
            arrow_key_up_selects_row=True,
            arrow_key_right_selects_column=True,
            arrow_key_left_selects_column=True,
            tab_key_right_selects_column=True,
            enter_key_down_selects_row=True,
            header_font=("微软雅黑", 9, "bold"),
            header_fg="black",
            header_bg="#f0f0f0",
            table_font=("微软雅黑", 9, "normal"),
            align="center",
            outline_thickness=0,  # 禁用边框以提升性能
            outline_color="#cccccc",
            show_default_header_for_empty=True,
            show_default_index_for_empty=True,
            show_top_left=True,
            top_left_bg="#f0f0f0",
            top_left_fg="black",
            header_height=25,
            row_index_width=0,
            column_index_width=0,
            drag_and_drop=False,
            double_click_column_resize=False,
            double_click_row_resize=False,
            right_click_selects_column=False,
            right_click_selects_row=False,
            right_click_deselects=False,
            shift_selects=True,
            ctrl_selects=True,
            drag_selects=True,
            drag_selection=True,
            select_all=True,
            deselect_all=True,
            copy=True,
            cut=False,
            paste=False,
            delete=False,
            undo=False,
            redo=False,
            insert_column=False,
            insert_row=False,
            delete_column=False,
            delete_row=False,
            move_rows=False,
            move_columns=False,
            merge_cells=False,
            unmerge_cells=False,
            resize_column=False,
            resize_row=False,
            rename_column=False,
            rename_row=False,
            hide_columns=False,
            hide_rows=False,
            freeze_columns=False,
            freeze_rows=False,
            freeze_header=False,
            freeze_index=False,
            freeze_top_left=False,
            freeze_top_left_columns=0,
            freeze_top_left_rows=0,
            freeze_header_rows=0,
            freeze_index_columns=0,
            show_scroll_bars=True,
            show_horizontal_scrollbar=False,  # 禁用横向滚动条
            show_vertical_scrollbar=True,
            scrollbar_width=15,
            scrollbar_color="#cccccc",
            scrollbar_troughcolor="#f0f0f0",
            scrollbar_button_color="#cccccc",
            scrollbar_button_hover_color="#999999",
            scrollbar_button_active_color="#666666",
            scrollbar_thumb_color="#999999",
            scrollbar_thumb_hover_color="#666666",
            scrollbar_thumb_active_color="#333333",
            scrollbar_border_color="#cccccc",
            scrollbar_border_width=1,
            scrollbar_border_style="solid",
            scrollbar_border_radius=0,
            scrollbar_corner_color="#f0f0f0",
            scrollbar_corner_border_color="#cccccc",
            scrollbar_corner_border_width=1,
            scrollbar_corner_border_style="solid",
            scrollbar_corner_border_radius=0,
        )
        
        sheet.pack(fill=tk.BOTH, expand=True)
        
        # 绑定双击事件
        def handle_double_click(event):
            try:
                selected = sheet.get_selected_rows()
                if selected:
                    row = selected[0]
                    df = sheet.get_sheet_data()
                    if row < len(df):
                        blank_code = df[row][2]  # 毛坯编码在第3列（索引2）
                        if blank_code and blank_code != "-":
                            show_blank_code_orders(blank_code, global_data)
            except Exception as e:
                pass
        
        sheet.bind("<Double-Button-1>", handle_double_click)
        
        widgets["table_analysis"] = sheet
        widgets["tree_analysis"] = sheet  # 向后兼容
        widgets["table_columns"] = table_columns
    else:
        # 回退到Treeview
        ttk.Label(frame_table, text="tksheet未安装，使用Treeview").pack()
        from ui_analysis import build_analysis_tab as build_treeview_tab
        return build_treeview_tab(parent, global_data, path_config)
    
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
            
            # 仅导出表格中显示的字段
            export_columns = [
                "成品物料编码", "成品物料名称", "毛坯物料编码", "毛坯物料名称",
                "年预计量", "精准季预计量", "订单数量", "成品库存数量", "毛坯库存数量", "已下单数量", "库存健康状态"
            ]
            export_df = analysis_df[export_columns].copy()
            
            # 重命名列
            export_df.rename(columns={
                "成品物料编码": "成品编码",
                "成品物料名称": "成品名称",
                "毛坯物料编码": "毛坯编码",
                "毛坯物料名称": "毛坯名称",
                "成品库存数量": "成品库存",
                "毛坯库存数量": "毛坯库存",
                "库存健康状态": "库存状态"
            }, inplace=True)
            
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
    更新分析结果表格数据（使用tksheet）
    """
    if not TKSHEET_AVAILABLE:
        return
    
    sheet = widgets["table_analysis"]
    
    # 获取分析数据
    analysis_df = global_data.get("analysis_show_df")
    if analysis_df is None or analysis_df.empty:
        sheet.data = [[""] * len(widgets["table_columns"])]
        return
    
    # 筛选数据
    if filter_type == "urgent":
        display_df = analysis_df[analysis_df["建议下单数量"] > 0].copy()
    else:
        display_df = analysis_df.copy()
    
    # 字段映射
    column_mapping = {
        "成品编码": "成品物料编码",
        "成品名称": "成品物料名称",
        "毛坯编码": "毛坯物料编码",
        "毛坯名称": "毛坯物料名称",
        "总订单": "总订单",
        "今日订单": "今日订单数量",
        "年预计量": "年预计量",
        "精准季预计量": "精准季预计量",
        "成品库存": "成品库存数量",
        "毛坯库存": "毛坯库存数量",
        "已下单数量": "已下单数量",
        "库存状态": "建议下单数量"
    }
    
    # 数据格式化
    numeric_fields = ["总订单", "今日订单数量", "年预计量", "精准季预计量", "成品库存数量", "毛坯库存数量", "已下单数量", "建议下单数量"]
    for field in numeric_fields:
        if field in display_df.columns:
            display_df[field] = pd.to_numeric(display_df[field], errors="coerce").fillna(0).astype(int)
    
    if "颜色状态" not in display_df.columns:
        display_df["颜色状态"] = "normal"
    
    # 构建显示数据
    table_data = []
    for idx in range(len(display_df)):
        row = display_df.iloc[idx]
        row_values = []
        for table_col in widgets["table_columns"]:
            analysis_col = column_mapping[table_col]
            val = row.get(analysis_col, "-")
            if pd.isna(val) or val == "":
                val = "-"
            elif isinstance(val, (int, float)):
                val = str(int(val))
            else:
                val = str(val)
            row_values.append(val)
        table_data.append(row_values)
    
    # 更新表格
    sheet.data = table_data
    
    # 应用背景色（基于颜色状态）- tksheet API 可能不支持这种方式，暂时禁用
    # if "颜色状态" in display_df.columns:
    #     for idx in range(len(display_df)):
    #         color_status = display_df.iloc[idx]["颜色状态"]
    #         if color_status == "red":
    #             sheet.row_index(idx, bg="#FFCCCC")
    #         elif color_status == "yellow":
    #             sheet.row_index(idx, bg="#FFFFCC")

def update_analysis_cards(widgets, global_data):
    """更新预警卡片数据（成品维度）"""
    analysis_df = global_data.get("analysis_show_df")
    if analysis_df is None or analysis_df.empty:
        widgets["card_need_purchase"].config(text="需下单：0件")
        widgets["card_need_warning"].config(text="需警惕：0件")
        widgets["card_normal"].config(text="正常：0件")
        return
    
    if "颜色状态" not in analysis_df.columns:
        analysis_df["颜色状态"] = "normal"
    
    urgent_count = len(analysis_df[analysis_df["颜色状态"] == "red"])
    warning_count = len(analysis_df[analysis_df["颜色状态"] == "yellow"])
    normal_count = len(analysis_df[analysis_df["颜色状态"] == "normal"])
    
    widgets["card_need_purchase"].config(text=f"需下单：{urgent_count}件")
    widgets["card_need_warning"].config(text=f"需警惕：{warning_count}件")
    widgets["card_normal"].config(text=f"正常：{normal_count}件")

def show_blank_code_orders(blank_code, global_data):
    """
    显示指定毛坯编码的订单记录（新窗口）
    """
    # 检查已下单文件是否存在
    ordered_df = global_data.get("ordered_df")
    if ordered_df is None or ordered_df.empty:
        messagebox.showwarning("提示", "未找到已下单文件数据！")
        return
    
    # 根据毛坯编码过滤数据
    filtered_df = ordered_df[ordered_df['毛坯编码'].str.strip() == blank_code.strip()].copy()
    
    if filtered_df.empty:
        messagebox.showinfo("提示", f"未找到毛坯编码【{blank_code}】的订单记录！")
        return
    
    # 创建新窗口
    new_window = tk.Toplevel()
    new_window.title(f"毛坯编码【{blank_code}】的订单记录")
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
    
    # 创建表格列（去除采购日期和成品编码列）
    exclude_columns = ['采购日期', '成品编码']
    table_columns = [col for col in ordered_df.columns if col not in exclude_columns]
    
    # 配置表格样式
    style = ttk.Style()
    style.configure("Order.Treeview", font=("微软雅黑", 9))
    style.configure("Order.Treeview.Heading", font=("微软雅黑", 9, "bold"))
    
    tree_orders = ttk.Treeview(new_window, columns=table_columns, show="headings", style="Order.Treeview")
    
    # 设置表头和列宽（所有列居中对齐）
    for col in table_columns:
        tree_orders.heading(col, text=col)
        tree_orders.column(col, width=100, minwidth=50, anchor='center')
    
    # 布局
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
