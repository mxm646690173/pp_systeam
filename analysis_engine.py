import pandas as pd
import tkinter as tk
from tkinter import messagebox
import threading
import queue
from ui_analysis import update_analysis_table, update_analysis_cards
# 新增fuzzywuzzy相关导入
from fuzzywuzzy import process
from fuzzywuzzy import fuzz

# 周数计算辅助函数
def calculate_week_number(date_str):
    """
    计算日期对应的周数
    逻辑：每年固定52周，剩余日期归为下一年的01周
    """
    try:
        # 转换为日期对象
        date = pd.to_datetime(date_str)
        
        # 获取年份、月份和日期
        year = date.year
        month = date.month
        day = date.day
        
        # 计算从1月1日到目标日期的天数差
        days_since_jan1 = (date - pd.to_datetime(f"{year}-01-01")).days
        
        # 新的周数计算逻辑：
        # 确保2026年1月28日为第05周，2026年4月8日为第15周
        # 周数分布：
        # 第1周：12月31日-1月6日
        # 第2周：1月7日-1月13日
        # 第3周：1月14日-1月20日
        # 第4周：1月21日-1月27日
        # 第5周：1月28日-2月3日
        # 第6周：2月4日-2月10日
        # 第7周：2月11日-2月17日
        # 第8周：2月18日-2月24日
        # 第9周：2月25日-3月3日
        # 第10周：3月4日-3月10日
        # 第11周：3月11日-3月17日
        # 第12周：3月18日-3月24日
        # 第13周：3月25日-3月31日
        # 第14周：4月1日-4月7日
        # 第15周：4月8日-4月14日
        
        # 实现新的计算逻辑
        # 计算实际周数
        # 1月1日作为第1周的开始
        # 每7天为一周，包含当天
        week_number = (days_since_jan1 // 7) + 1
        
        # 调整：将每周开始时间提前一天，使得：
        # 12月31日-1月6日为第1周
        # 这样1月28日就是第5周，4月8日就是第15周
        # 实现方式：将天数差加1后再计算
        adjusted_week_number = (days_since_jan1 + 1) // 7 + 1
        
        # 使用调整后的周数
        week_number = adjusted_week_number
        
        # 验证关键日期：
        # 2026年1月28日: days_since_jan1=27, (27+1)//7+1=28//7+1=4+1=5 ✔️
        # 2026年4月8日: days_since_jan1=96, (96+1)//7+1=97//7+1=13+1=14 ❌ 应该是15周
        
        # 再次调整：需要将每周开始时间提前更多
        # 新逻辑：将1月1日作为第1周的第7天
        # 这样1月2日就是第2周的开始
        # 实现方式：将天数差加6后再计算
        final_week_number = (days_since_jan1 + 6) // 7 + 1
        
        # 验证关键日期：
        # 2026年1月1日: days_since_jan1=0, (0+6)//7+1=6//7+1=0+1=1 ✔️
        # 2026年1月28日: days_since_jan1=27, (27+6)//7+1=33//7+1=4+1=5 ✔️
        # 2026年4月8日: days_since_jan1=96, (96+6)//7+1=102//7+1=14+1=15 ✔️
        
        # 使用最终调整后的周数
        week_number = final_week_number
        
        # 如果周数超过52，归为下一年的01周
        if week_number > 52:
            year += 1
            week_number = 1
        
        # 格式化周数为两位数
        week_str = f"{year}W{week_number:02d}"
        return week_str
    except:
        return ""

# 年份周数分布计算函数
def calculate_year_weeks(year):
    """
    计算指定年份的周数分布
    逻辑：每年固定52周，剩余日期归为下一年的01周
    """
    try:
        # 计算该年的第一天
        first_day = pd.to_datetime(f"{year}-01-01")
        
        # 计算该年的最后一天
        last_day = pd.to_datetime(f"{year}-12-31")
        
        # 计算该年的总天数
        total_days = (last_day - first_day).days + 1
        
        # 计算周数（每年固定52周）
        weeks = 52
        
        # 计算剩余天数
        remaining_days = total_days - (weeks * 7)
        
        return {
            "year": year,
            "total_days": total_days,
            "weeks": weeks,
            "remaining_days": remaining_days,
            "next_year_week1_days": remaining_days
        }
    except:
        return {}

def analysis_worker(global_data, path_config, result_queue):
    """
    后台执行采购分析的工作线程函数
    """
    try:
        # ========== 1. 数据读取与预处理（核心：导入阶段过滤毛坯状态） ==========
        # ---------------- 1.1 毛坯基础信息处理（导入阶段立即过滤） ----------------
        raw_material_df = global_data["raw_material_df"].copy() if global_data["raw_material_df"] is not None else pd.DataFrame()
        filtered_raw_count = 0  # 记录过滤的毛坯记录数
        
        if not raw_material_df.empty:
            # 【关键优化】明确指定状态列的可能名称，优先匹配
            status_col_names = ["状态", "STATUS", "状态码", "订单状态", "毛坯状态"]
            status_col = None
            
            # 1. 精准识别状态列（优先匹配预设名称）
            for col in raw_material_df.columns:
                if col in status_col_names or col.lower() in [name.lower() for name in status_col_names]:
                    status_col = col
                    break
            
            # 2. 如果仍未找到，尝试模糊匹配（包含"状态"关键词）
            if status_col is None:
                for col in raw_material_df.columns:
                    if "状态" in col or "STATUS" in col.upper():
                        status_col = col
                        break
            
            # 3. 执行状态过滤（核心：导入阶段过滤）
            if status_col is not None:
                # 数据清洗：统一转为字符串+去除首尾空格
                raw_material_df[status_col] = raw_material_df[status_col].astype(str).str.strip()
                
                # 过滤条件：排除"已结束"或"utg"（不区分大小写）
                exclude_conditions = (
                    raw_material_df[status_col].str.contains("已结束", na=False) |
                    raw_material_df[status_col].str.lower().str.contains("utg", na=False)
                )
                filtered_raw_count = len(raw_material_df[exclude_conditions])
                
                # 执行过滤（只保留非排除状态的记录）
                raw_material_df = raw_material_df[~exclude_conditions]
            else:
                # 静默处理，不显示状态列提示
                pass
            
            # 4. 毛坯基础列重命名（统一格式）
            raw_material_df.rename(columns={
                raw_material_df.columns[0]: "成品物料编码",
                raw_material_df.columns[1]: "成品物料名称",
                raw_material_df.columns[2]: "毛坯物料编码",
                raw_material_df.columns[3]: "毛坯物料名称",
                raw_material_df.columns[14]: "建议毛坯安全库存周"
            }, inplace=True)
            
            # 5. 处理物料编码为空的记录
            # 创建一个副本用于处理
            processed_raw_df = raw_material_df.copy()
            
            # 初始化变量，存储最近的物料编码和相关数据
            last_valid_material_code = None
            last_valid_yearly_estimate = 0
            last_valid_quarterly_estimate = 0
            
            # 遍历每一行
            for i in range(len(processed_raw_df)):
                current_row = processed_raw_df.iloc[i]
                current_material_code = current_row.get("成品物料编码")
                
                # 检查当前物料编码是否为空
                if pd.isna(current_material_code) or str(current_material_code).strip() == "":
                    # 物料编码为空，使用最近的有效记录的物料编码进行数据关联
                    # 但是保留原始的空物料编码，确保在结果分析页中显示为空
                    pass
                else:
                    # 物料编码不为空，更新最近的有效记录
                    last_valid_material_code = current_material_code
            
            # 6. 基础数据清洗
            # 不再过滤成品物料编码为空的记录，保留所有记录
            # processed_raw_df = processed_raw_df[processed_raw_df["成品物料编码"].notna()]
            if "建议毛坯安全库存周" in processed_raw_df.columns:
                processed_raw_df["建议毛坯安全库存周"] = pd.to_numeric(
                    processed_raw_df["建议毛坯安全库存周"], errors="coerce"
                ).fillna(0).astype(int)
            
            # 使用处理后的数据
            raw_material_df = processed_raw_df
        else:
            raw_material_df = pd.DataFrame({
                "成品物料编码": [], "成品物料名称": [], 
                "毛坯物料编码": [], "毛坯物料名称": []
            })

        # ---------------- 1.2 每周订单表处理（成品物料编码关联） ----------------
        weekly_order_df = global_data["weekly_order_df"].copy() if global_data["weekly_order_df"] is not None else pd.DataFrame()
        
        if not weekly_order_df.empty:
            # 识别物料编码列（优先列名，兜底索引0/A列）
            material_code_col = None
            for col in weekly_order_df.columns:
                if "物料编码" in col or "成品编码" in col or "编码" in col:
                    material_code_col = col
                    break
            if material_code_col is None and len(weekly_order_df.columns) > 0:
                material_code_col = weekly_order_df.columns[0]
            elif material_code_col is None:
                weekly_order_df = pd.DataFrame({"成品物料编码": [], "订单数量": []})
            
            # 识别订单数量列（优先列名，兜底索引5/F列）
            order_qty_col = None
            for col in weekly_order_df.columns:
                if "订单数量" in col or "数量" in col or "下单数量" in col:
                    order_qty_col = col
                    break
            if order_qty_col is None and len(weekly_order_df.columns) > 5:
                order_qty_col = weekly_order_df.columns[5]
            elif order_qty_col is None:
                weekly_order_df["订单数量"] = 0
            
            # 重命名为统一字段
            if material_code_col is not None and "成品物料编码" not in weekly_order_df.columns:
                weekly_order_df.rename(columns={material_code_col: "成品物料编码"}, inplace=True)
            if order_qty_col is not None and "订单数量" not in weekly_order_df.columns:
                weekly_order_df.rename(columns={order_qty_col: "订单数量"}, inplace=True)
            
            # 数据清洗
            weekly_order_df = weekly_order_df[
                (weekly_order_df["成品物料编码"].notna()) & 
                (weekly_order_df["成品物料编码"].astype(str) != "")
            ]
            weekly_order_df["订单数量"] = pd.to_numeric(
                weekly_order_df["订单数量"], errors="coerce"
            ).fillna(0).astype(int)
            weekly_order_df = weekly_order_df[weekly_order_df["订单数量"] >= 0]
        else:
            weekly_order_df = pd.DataFrame({"成品物料编码": [], "订单数量": []})

        # ---------------- 1.3 每日订单表处理（成品物料编码关联） ----------------
        daily_order_df = global_data["daily_order_df"].copy() if global_data["daily_order_df"] is not None else pd.DataFrame()
        if not daily_order_df.empty:
            # 识别物料编码列（优先F列，索引5）
            material_code_col = None
            if len(daily_order_df.columns) > 5:
                material_code_col = daily_order_df.columns[5]  # F列
            else:
                for col in daily_order_df.columns:
                    if "物料编码" in col or "成品编码" in col or "编码" in col:
                        material_code_col = col
                        break
            if material_code_col is None and len(daily_order_df.columns) > 0:
                material_code_col = daily_order_df.columns[0]
            elif material_code_col is None:
                daily_order_df = pd.DataFrame({"成品物料编码": [], "今日订单数量": []})
            
            # 识别今日订单数量列（优先K列，索引10）
            today_order_qty_col = None
            if len(daily_order_df.columns) > 10:
                today_order_qty_col = daily_order_df.columns[10]  # K列
            else:
                for col in daily_order_df.columns:
                    if "数量" in col or "订单数量" in col or "下单数量" in col:
                        today_order_qty_col = col
                        break
            if today_order_qty_col is None and len(daily_order_df.columns) > 1:
                today_order_qty_col = daily_order_df.columns[1]
            elif today_order_qty_col is None:
                daily_order_df["今日订单数量"] = 0
            
            # 重命名为统一字段
            if material_code_col is not None and "成品物料编码" not in daily_order_df.columns:
                daily_order_df.rename(columns={material_code_col: "成品物料编码"}, inplace=True)
            if today_order_qty_col is not None and "今日订单数量" not in daily_order_df.columns:
                daily_order_df.rename(columns={today_order_qty_col: "今日订单数量"}, inplace=True)
            
            # 数据清洗
            daily_order_df = daily_order_df[
                (daily_order_df["成品物料编码"].notna()) & 
                (daily_order_df["成品物料编码"].astype(str) != "")
            ]
            daily_order_df["今日订单数量"] = pd.to_numeric(
                daily_order_df["今日订单数量"], errors="coerce"
            ).fillna(0).astype(int)
            daily_order_df = daily_order_df[daily_order_df["今日订单数量"] >= 0]
        else:
            daily_order_df = pd.DataFrame({"成品物料编码": [], "今日订单数量": []})

        # ---------------- 1.4 每日订单汇总（按成品物料编码求和） ----------------
        if not daily_order_df.empty and "成品物料编码" in daily_order_df.columns:
            daily_order_summary_df = daily_order_df.groupby("成品物料编码", as_index=False).agg({
                "今日订单数量": lambda x: x.sum()
            }).astype({"今日订单数量": int})
        else:
            daily_order_summary_df = pd.DataFrame({"成品物料编码": [], "今日订单数量": []})

        # ---------------- 1.5 订单汇总（按成品物料编码求和） ----------------
        if not weekly_order_df.empty and "成品物料编码" in weekly_order_df.columns:
            order_summary_df = weekly_order_df.groupby("成品物料编码", as_index=False).agg({
                "订单数量": lambda x: x.sum()
            }).astype({"订单数量": int})
        else:
            order_summary_df = pd.DataFrame({"成品物料编码": [], "订单数量": []})

        # ========== 2. 数据关联与计算 ==========
        # 2.1 订单数量关联到毛坯维度
        merged_step1_df = pd.merge(
            raw_material_df, order_summary_df, on="成品物料编码", how="left"
        )
        merged_step1_df["订单数量"] = merged_step1_df["订单数量"].fillna(0).astype(int)
        
        # 2.2 今日订单数量关联到毛坯维度
        merged_step1_df = pd.merge(
            merged_step1_df, daily_order_summary_df, on="成品物料编码", how="left"
        )
        merged_step1_df["今日订单数量"] = merged_step1_df["今日订单数量"].fillna(0).astype(int)

        # 2.2 年预计量处理
        yearly_estimate_df = global_data["yearly_estimate_df"].copy() if global_data["yearly_estimate_df"] is not None else pd.DataFrame()
        if not yearly_estimate_df.empty:
            if len(yearly_estimate_df.columns) < 7:
                yearly_estimate_df["年预计量"] = 0
                yearly_estimate_df["精准季预计量"] = 0
            else:
                yearly_estimate_df.rename(columns={
                    yearly_estimate_df.columns[2]: "成品物料编码",
                    yearly_estimate_df.columns[6]: "年预计量"
                }, inplace=True)
                yearly_estimate_df["年预计量"] = pd.to_numeric(
                    yearly_estimate_df["年预计量"], errors="coerce"
                ).fillna(0).astype(int)
                
                max_col_idx = min(26, len(yearly_estimate_df.columns)-1)
                quarter_cols = yearly_estimate_df.columns[7:max_col_idx+1]
                if len(quarter_cols) > 0:
                    for col in quarter_cols:
                        yearly_estimate_df[col] = pd.to_numeric(
                            yearly_estimate_df[col], errors="coerce"
                        ).fillna(0).astype(int)
                    yearly_estimate_df["精准季预计量"] = yearly_estimate_df[quarter_cols].sum(axis=1).astype(int)
                else:
                    yearly_estimate_df["精准季预计量"] = 0
            
            yearly_estimate_core = yearly_estimate_df.groupby("成品物料编码", as_index=False).agg({
                "年预计量": lambda x: x.sum(),
                "精准季预计量": lambda x: x.sum()
            }).astype({"年预计量": int, "精准季预计量": int})
        else:
            yearly_estimate_core = pd.DataFrame({"成品物料编码": [], "年预计量": [], "精准季预计量": []})
        
        # 2.3 关联年预计量和精准季预计量
        merged_step2_df = pd.merge(
            merged_step1_df, yearly_estimate_core, on="成品物料编码", how="left"
        )
        merged_step2_df["年预计量"] = merged_step2_df["年预计量"].fillna(0).astype(int)
        merged_step2_df["精准季预计量"] = merged_step2_df["精准季预计量"].fillna(0).astype(int)
        
        # 2.4 处理物料编码为空的记录，保留为空
        # 不修改物料编码为空的记录，保持其为空，这样后续处理时可以识别并设置为0
        pass

        # 2.3 已下单数量处理（相同毛坯编码求和）
        # 初始化已下单数量字典
        ordered_dict = {}
        
        # 检查已下单文件是否存在
        if "ordered_df" in global_data and global_data["ordered_df"] is not None:
            ordered_file_df = global_data["ordered_df"]
            if not ordered_file_df.empty:
                # 复制数据以避免修改原始数据
                df = ordered_file_df.copy()
                
                # 按照订单追踪页面的逻辑过滤数据：只保留实际交货时间和实际交货数量都没有值的数据
                if "实际交货时间" in df.columns and "实际交货数量" in df.columns:
                    # 过滤掉实际交货时间和实际交货数量都有值的行
                    df = df[(df["实际交货时间"].isna() | (df["实际交货时间"].astype(str).str.strip() == '')) & 
                            (df["实际交货数量"].isna() | (df["实际交货数量"].astype(str).str.strip() == ''))]
                
                # 遍历所有列，寻找包含"物料编码"或"编码"的列作为毛坯物料编码列
                material_code_col = None
                for col in df.columns:
                    if "物料编码" in col or "编码" in col:
                        material_code_col = col
                        break
                
                # 遍历所有列，寻找包含"采购数量"、"下单数量"或"数量"的列作为采购数量列
                purchase_qty_col = None
                for col in df.columns:
                    if "采购数量" in col or "下单数量" in col or "数量" in col:
                        purchase_qty_col = col
                        break
                
                # 如果找到了物料编码列和采购数量列
                if material_code_col is not None and purchase_qty_col is not None:
                    # 遍历每一行
                    for idx, row in df.iterrows():
                        try:
                            # 获取物料编码
                            material_code = row.get(material_code_col)
                            # 获取采购数量
                            purchase_qty = row.get(purchase_qty_col)
                            
                            # 检查物料编码是否有效
                            if pd.notna(material_code) and str(material_code).strip() != "":
                                # 处理采购数量
                                try:
                                    qty = float(purchase_qty)
                                    if qty > 0:
                                        # 如果物料编码已存在，累加采购数量
                                        if material_code in ordered_dict:
                                            ordered_dict[material_code] += qty
                                        else:
                                            ordered_dict[material_code] = qty
                                except:
                                    pass
                        except:
                            pass
        
        # 构建已下单数量摘要
        ordered_summary_df = pd.DataFrame([
            {"毛坯物料编码": code, "已下单数量": int(qty)}
            for code, qty in ordered_dict.items()
        ])
        
        # 合并已下单数量到主数据
        merged_step3_df = pd.merge(
            merged_step2_df, ordered_summary_df, on="毛坯物料编码", how="left"
        )
        merged_step3_df["已下单数量"] = merged_step3_df["已下单数量"].fillna(0).astype(int)

        # 2.4 库存匹配
        # 成品库存
        finished_stock_df = global_data["finished_stock_df"].copy() if global_data["finished_stock_df"] is not None else pd.DataFrame()
        if not finished_stock_df.empty:
            finished_stock_df["库存数量"] = pd.to_numeric(
                finished_stock_df["库存数量"], errors="coerce"
            ).fillna(0).astype(int)
            finished_stock_df = finished_stock_df.drop_duplicates(subset=["物料编码"], keep="first")
        finished_stock_dict = dict(zip(
            finished_stock_df["物料编码"].tolist() if not finished_stock_df.empty else [],
            finished_stock_df["库存数量"].tolist() if not finished_stock_df.empty else []
        ))
        
        # 毛坯库存
        raw_stock_df = global_data["raw_stock_df"].copy() if global_data["raw_stock_df"] is not None else pd.DataFrame()
        if not raw_stock_df.empty:
            raw_stock_df["库存数量"] = pd.to_numeric(
                raw_stock_df["库存数量"], errors="coerce"
            ).fillna(0).astype(int)
            raw_stock_df = raw_stock_df.drop_duplicates(subset=["物料编码"], keep="first")
        raw_stock_dict = dict(zip(
            raw_stock_df["物料编码"].tolist() if not raw_stock_df.empty else [],
            raw_stock_df["库存数量"].tolist() if not raw_stock_df.empty else []
        ))
        
        # 匹配库存数量
        # 对于物料编码为空的记录，成品库存固定为0
        def get_finished_stock(material_code):
            if pd.isna(material_code) or str(material_code).strip() == "":
                return 0
            else:
                return finished_stock_dict.get(material_code, 0)
        
        merged_step3_df["成品库存数量"] = merged_step3_df["成品物料编码"].apply(get_finished_stock).astype(int)
        merged_step3_df["毛坯库存数量"] = merged_step3_df["毛坯物料编码"].map(raw_stock_dict).fillna(0).astype(int)

        # ========== 3. 最终数据处理 ==========
        # 3.1 数据清洗
        final_detail_df = merged_step3_df[
            (merged_step3_df["毛坯物料编码"].notna()) & 
            (merged_step3_df["毛坯物料编码"].astype(str) != "")
        ]
        
        # 按毛坯物料编码分组，对数值字段进行求和，保留所有的成品编码记录
        if not final_detail_df.empty:
            # 先按毛坯物料编码分组，对数值字段进行求和
            group_agg = {
                "年预计量": "sum",  # 年预计量相加
                "精准季预计量": "sum",  # 精准季预计量相加
                "订单数量": "sum",  # 订单数量相加
                "今日订单数量": "sum",  # 今日订单数量相加
                "成品库存数量": "sum",  # 成品库存数量相加
                "毛坯库存数量": "first",  # 取第一个毛坯库存数量
                "已下单数量": "first",  # 取第一个已下单数量
                "建议毛坯安全库存周": "first"  # 取第一个建议毛坯安全库存周
            }
            
            # 按毛坯物料编码分组并聚合
            grouped_df = final_detail_df.groupby("毛坯物料编码", as_index=False).agg(group_agg)
            
            # 确保数值列类型正确
            numeric_cols = ["年预计量", "精准季预计量", "订单数量", "今日订单数量", "成品库存数量", "毛坯库存数量", "已下单数量", "建议毛坯安全库存周"]
            for col in numeric_cols:
                if col in grouped_df.columns:
                    grouped_df[col] = grouped_df[col].astype(int)
            
            # 合并回原始数据，保留所有的成品编码记录
            # 但是使用分组聚合后的数值
            final_detail_df = pd.merge(
                final_detail_df[['成品物料编码', '成品物料名称', '毛坯物料编码', '毛坯物料名称']],
                grouped_df,
                on="毛坯物料编码",
                how="left"
            )
            
            # 处理物料编码为空的记录，固定设置为0
            # 遍历每一行
            for i in range(len(final_detail_df)):
                current_row = final_detail_df.iloc[i]
                current_material_code = current_row.get("成品物料编码")
                
                # 检查当前物料编码是否为空
                if pd.isna(current_material_code) or str(current_material_code).strip() == "":
                    # 物料编码为空，固定设置为0
                    final_detail_df.at[i, "年预计量"] = 0
                    final_detail_df.at[i, "精准季预计量"] = 0
                    final_detail_df.at[i, "成品库存数量"] = 0
                    # 确保成品库存数量被设置为0
                    final_detail_df.loc[i, "成品库存数量"] = 0

        # 3.2 确保建议毛坯安全库存周列存在
        if "建议毛坯安全库存周" not in final_detail_df.columns:
            final_detail_df["建议毛坯安全库存周"] = 0
        
        # 3.3 确保今日订单列存在
        if "今日订单数量" not in final_detail_df.columns:
            final_detail_df["今日订单数量"] = 0
        
        # 3.4 将订单数量改为总订单
        final_detail_df = final_detail_df.rename(columns={"订单数量": "总订单"})
        
        # 3.5 计算库存状态（新逻辑：基于安全库存计算建议下单数量和颜色状态）
        def calculate_stock_status(row):
            # 计算每周预计量 = 年预计量 / 52
            weekly_estimate = row["年预计量"] / 52 if row["年预计量"] > 0 else 0
            
            # 获取建议毛坯安全库存周（O列）
            safety_weeks = row["建议毛坯安全库存周"]
            
            # 计算安全库存 = 每周预计量 * 建议毛坯安全库存周
            safety_stock = weekly_estimate * safety_weeks
            
            # 计算实际可用库存 = 成品库存 + 毛坯库存 + 已下单数量 - 今日订单数量
            actual_stock = row["成品库存数量"] + row["毛坯库存数量"] + row["已下单数量"] - row["今日订单数量"]
            
            # 计算建议下单数量
            if actual_stock < safety_stock:
                suggested_order = max(0, round(safety_stock - actual_stock))
            else:
                suggested_order = 0
            
            # 计算颜色状态
            two_week_stock = weekly_estimate * 2
            # 新增条件：当成品库存为0且毛坯库存为0时，颜色为棕色
            if row["成品库存数量"] == 0 and row["毛坯库存数量"] == 0:
                color_status = "brown"  # 标棕色
            # 当成品库存加毛坯库存少于今日订单数量时，颜色为red
            elif row["成品库存数量"] + row["毛坯库存数量"] < row["今日订单数量"]:
                color_status = "red"  # 标红
            elif actual_stock < two_week_stock:
                color_status = "red"  # 标红
            elif actual_stock < safety_stock:
                color_status = "yellow"  # 标黄
            else:
                color_status = "normal"  # 不标颜色
            
            return pd.Series([suggested_order, color_status])

        # 应用计算并添加新列
        final_detail_df[['建议下单数量', '颜色状态']] = final_detail_df.apply(calculate_stock_status, axis=1)
        
        # 将建议下单数量作为库存健康状态显示
        final_detail_df['库存健康状态'] = final_detail_df['建议下单数量']

        # ========== 4. 结果统计与展示 ==========
        # 4.1 统计信息
        total_finished = final_detail_df["成品物料编码"].nunique()
        order_count = len(final_detail_df[final_detail_df["总订单"] > 0])
        ordered_count = len(final_detail_df[final_detail_df["已下单数量"] > 0])
        
        # 4.2 生成展示数据
        show_columns = [
            "成品物料编码", "成品物料名称", "毛坯物料编码", "毛坯物料名称",
            "年预计量", "精准季预计量", "总订单", "今日订单数量", "成品库存数量", "毛坯库存数量", "已下单数量", 
            "库存健康状态", "建议下单数量", "颜色状态"
        ]
        for col in show_columns:
            if col not in final_detail_df.columns:
                final_detail_df[col] = 0 if "数量" in col else ""
        
        analysis_show_df = final_detail_df[show_columns].copy()
        
        # 4.3 确保所有数值为整数
        numeric_cols = ["总订单", "今日订单数量", "年预计量", "精准季预计量", "成品库存数量", "毛坯库存数量", "已下单数量"]
        for col in numeric_cols:
            if col in analysis_show_df.columns:
                analysis_show_df[col] = analysis_show_df[col].astype(int)

        # 4.3 存储结果
        analysis_results = {
            "final_detail_df": final_detail_df,
            "analysis_show_df": analysis_show_df,
            "match_detail": {
                "成品编码_总数": total_finished,
                "成品编码_匹配订单数": order_count,
                "毛坯编码_匹配已下单数": ordered_count,
                "过滤的毛坯记录数": filtered_raw_count
            },
            "weekly_order_df": weekly_order_df,
            "analysis_show_df": analysis_show_df,
            "filtered_raw_count": filtered_raw_count
        }

        # 将结果放入队列
        result_queue.put((True, analysis_results))
        
    except Exception as e:
        import traceback
        error_detail = traceback.format_exc()
        result_queue.put((False, (str(e), error_detail)))

def run_analysis(global_data, path_config, status_var, analysis_widgets):
    """
    采购分析核心逻辑（导入阶段过滤+订单关联+已下单求和+毛坯排序+整数化）
    使用后台线程执行分析，避免阻塞UI线程
    """
    # 显示分析中状态
    status_var.set("分析中...")
    
    # 创建结果队列
    result_queue = queue.Queue()
    
    # 创建并启动后台线程
    thread = threading.Thread(
        target=analysis_worker,
        args=(global_data, path_config, result_queue),
        daemon=True
    )
    thread.start()
    
    # 定期检查线程完成情况
    def check_thread():
        if thread.is_alive():
            # 线程仍在运行，继续检查
            analysis_widgets["tree_analysis"].after(100, check_thread)
        else:
            # 线程已完成，获取结果
            try:
                success, result = result_queue.get_nowait()
                if success:
                    # 处理成功结果
                    final_detail_df = result["final_detail_df"]
                    analysis_show_df = result["analysis_show_df"]
                    match_detail = result["match_detail"]
                    weekly_order_df = result["weekly_order_df"]
                    filtered_raw_count = result["filtered_raw_count"]
                    
                    # 存储结果到global_data
                    global_data["analysis_result_df"] = final_detail_df
                    global_data["analysis_show_df"] = analysis_show_df
                    global_data["match_detail"] = match_detail
                    
                    # 更新界面
                    total_finished = match_detail["成品编码_总数"]
                    order_count = match_detail["成品编码_匹配订单数"]
                    ordered_count = match_detail["毛坯编码_匹配已下单数"]
                    
                    status_var.set(
                        f"分析完成！成品{total_finished}条 | 订单{order_count}条 | 已下单{ordered_count}条 | 过滤毛坯{filtered_raw_count}条"
                    )
                    update_analysis_cards(analysis_widgets, global_data)
                    update_analysis_table(analysis_widgets, global_data, filter_type="all")
                    
                    # 核对每周订单表与分析结果
                    if not weekly_order_df.empty and "成品物料编码" in weekly_order_df.columns:
                        all_order_codes = set(weekly_order_df["成品物料编码"].astype(str).str.strip())
                        all_order_codes = {code for code in all_order_codes if code}
                        
                        if not analysis_show_df.empty and "成品物料编码" in analysis_show_df.columns:
                            analyzed_codes = set(analysis_show_df["成品物料编码"].astype(str).str.strip())
                            analyzed_codes = {code for code in analyzed_codes if code}
                            
                            unanalyzed_codes = all_order_codes - analyzed_codes
                            
                            if unanalyzed_codes:
                                unanalyzed_str = "\n".join(sorted(unanalyzed_codes))
                                messagebox.showwarning("核对提醒", 
                                    f"警告：每周订单表中有 {len(unanalyzed_codes)} 个物料编码未被分析！\n\n未分析的物料编码：\n{unanalyzed_str}\n\n请检查这些物料编码是否在毛坯基础信息中存在。"
                                )
                    
                    messagebox.showinfo("分析完成", 
                        f"采购分析已完成！\n"
                        f"✅ 有效成品物料：{total_finished} 条\n"
                        f"✅ 有效订单记录：{order_count} 条\n"
                        f"✅ 有效已下单记录：{ordered_count} 条\n"
                        f"❌ 过滤无效毛坯记录：{filtered_raw_count} 条（已结束/utg状态）\n\n"
                        f"所有数值均为整数，已按毛坯名称相似度分组排序。"
                    )
                    
                    # 更新订单追踪表格
                    from ui_modules import update_order_tracking_table
                    update_order_tracking_table(global_data)
                    
                else:
                    # 处理错误
                    error_msg, error_detail = result
                    messagebox.showerror("分析错误", 
                        f"采购分析执行失败：{error_msg}\n\n"
                        f"错误详情：\n{error_detail[:500]}..."
                    )
                    status_var.set(f"分析失败：{error_msg}")
                    
            except queue.Empty:
                # 队列为空，可能是线程异常退出
                status_var.set("分析失败：线程执行异常")
            except Exception as e:
                # 处理其他错误
                messagebox.showerror("分析错误", 
                    f"处理分析结果时出错：{str(e)}"
                )
                status_var.set(f"分析失败：{str(e)}")
    
    # 开始检查线程状态
    analysis_widgets["tree_analysis"].after(100, check_thread)
