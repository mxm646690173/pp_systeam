from tkinter import filedialog, messagebox, Toplevel, Listbox, LEFT, BOTH, Y, END
import tkinter as tk
from tkinter import ttk
import pandas as pd
import os
import datetime
import json
import requests
import urllib3

# 禁用SSL警告（如果需要）
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def init_data_folders(path_config):
    """初始化默认文件夹"""
    for folder in path_config.values():
        if not os.path.exists(folder) and not folder.endswith('.json'):
            os.makedirs(folder)

def download_weekly_order_file():
    """
    自动下载每周订单文件
    返回下载的文件路径，如果下载失败返回None
    """
    try:
        # 构建请求头
        headers = {
            'accept': 'application/json, text/plain, */*',
            'accept-encoding': 'gzip, deflate',
            'accept-language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
            'connection': 'keep-alive',
            'cookie': 'token=ba777f9b9a0940cbaa1a6abcc9cb8ac5',
            'host': '61.153.238.198:9331',
            'referer': 'http://61.153.238.198:9331/',
            'token': 'ba777f9b9a0940cbaa1a6abcc9cb8ac5',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/143.0.0.0 Safari/537.36 Edg/143.0.0.0'
        }
        
        # 生成当前时间戳
        import time
        timestamp = int(time.time() * 1000)
        
        # 构建请求URL
        url = f'http://61.153.238.198:9331/api-zsdae/bu/jrgyspurorder/exportExcel'
        params = {
            't': timestamp,
            'page': 1,
            'limit': 2000,
            'key': '',
            'fentryStatus': 0,
            'supinfo': '',
            'khddh': '',
            'furl': 'http://61.153.238.198:9331/print-zsdae/api/export/excelExport'
        }
        
        # 发送请求
        response = requests.get(url, headers=headers, params=params, verify=False)
        
        # 检查响应状态
        if response.status_code == 200:
            # 创建保存目录
            save_dir = './每周订单自动下载'
            if not os.path.exists(save_dir):
                os.makedirs(save_dir)
            
            # 生成文件名
            today = datetime.datetime.now().strftime('%Y%m%d')
            file_name = f'{today}_采购订单.xlsx'
            file_path = os.path.join(save_dir, file_name)
            
            # 尝试解析JSON响应
            try:
                json_data = response.json()
                
                # 检查JSON结构，根据实际返回的结构处理
                if isinstance(json_data, dict) and 'data' in json_data:
                    data = json_data['data']
                    
                    # 根据实际返回的数据结构，将各个字段的数据整合为DataFrame
                    if isinstance(data, dict):
                        # 直接构建数据字典，每个字段作为一列
                        data_dict = {}
                        
                        # 遍历所有字段
                        for field_name, field_values in data.items():
                            if isinstance(field_values, list):
                                # 如果是列表，直接使用
                                data_dict[field_name] = field_values
                            else:
                                # 如果不是列表，创建单个元素的列表
                                data_dict[field_name] = [field_values]
                        
                        # 转换为DataFrame
                        df = pd.DataFrame(data_dict)
                    else:
                        # 如果data不是字典，尝试直接转换
                        if isinstance(data, list):
                            df = pd.DataFrame(data)
                        else:
                            df = pd.DataFrame([data])
                else:
                    # 其他情况，尝试直接转换
                    if isinstance(json_data, list):
                        df = pd.DataFrame(json_data)
                    else:
                        df = pd.DataFrame([json_data])
                
                # 保存为Excel文件
                df.to_excel(file_path, index=False, sheet_name="喷涂订单整理")
            except Exception as json_e:
                # 如果JSON解析失败，尝试直接保存响应内容
                # print(f"JSON解析失败，尝试直接保存：{str(json_e)}")
                with open(file_path, 'wb') as f:
                    f.write(response.content)
            
            return file_path
        else:
            # print(f"下载失败，状态码: {response.status_code}")
            # print(f"响应内容: {response.text[:500]}...")  # 打印前500个字符查看错误
            return None
    except Exception as e:
        # print(f"下载过程中发生错误: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

def load_history(path_config):
    """
    从历史文件加载上次导入的文件路径
    """
    history_file = path_config.get("history", "./history.json")
    if os.path.exists(history_file):
        try:
            with open(history_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            # # print(f"加载历史记录失败：{str(e)}")
            pass
    return {}

def save_history(global_data, path_config):
    """
    保存当前的导入路径到历史文件
    """
    history_file = path_config.get("history", "./history.json")
    last_import_paths = global_data.get("last_import_paths", {})
    try:
        with open(history_file, 'w', encoding='utf-8') as f:
            json.dump(last_import_paths, f, ensure_ascii=False, indent=2)
    except Exception as e:
        # # print(f"保存历史记录失败：{str(e)}")
        pass

# ========== 新增：拆分后的毛坯基础信息导入函数 ==========
def import_raw_material_manual(global_data, status_var):
    """手动选择毛坯基础信息文件（固定M列过滤已结束/utg状态）"""
    raw_material_path = filedialog.askopenfilename(
        title="选择毛坯基础信息表",
        filetypes=( ("Excel文件", "*.xlsx"), ("所有文件", "*.*")))
    if not raw_material_path:
        return
    
    try:
        # 保存上次导入的路径
        global_data.setdefault("last_import_paths", {})["raw_material"] = raw_material_path
        
        # 1. 读取Excel文件（保留原始数据类型，M列后续单独处理）
        raw_df = pd.read_excel(raw_material_path)
        initial_count = len(raw_df)
        
        if initial_count == 0:
            messagebox.showwarning("警告", "毛坯基础信息表为空！")
            status_var.set("毛坯基础信息导入失败 - 表格为空")
            return
        
        # ========== 核心：固定M列（索引12）过滤无效状态 ==========
        # 确认M列（索引12）存在
        if len(raw_df.columns) <= 12:
            messagebox.showerror("错误", "毛坯基础信息表列数不足！M列（第13列）不存在")
            status_var.set("毛坯基础信息导入失败 - 无M列")
            return
        
        # 2. 取M列（索引12）作为状态列，统一处理格式
        status_col_index = 12  # M列 = 索引12（A=0, B=1...M=12）
        status_col_name = raw_df.columns[status_col_index]
        raw_df[status_col_name] = raw_df[status_col_name].astype(str).str.strip().str.lower()
        
        # 3. 执行过滤：排除"已结束" 或 "utg"
        valid_mask = ~(
            raw_df[status_col_name].str.contains("已结束", na=False) |
            raw_df[status_col_name].str.contains("utg", na=False)
        )
        filtered_df = raw_df[valid_mask].copy()
        filtered_count = initial_count - len(filtered_df)
        
        # 4. 存入全局数据（过滤后的有效数据）
        global_data["raw_material_df"] = filtered_df
        
        # 5. 更新界面和提示
        update_overview_table(global_data["tree_overview"], "毛坯基础信息", len(filtered_df))
        
        # 弹窗提示过滤结果
        if filtered_count > 0:
            messagebox.showinfo("导入成功", 
                f"毛坯基础信息导入成功！\n"
                f"📊 数据统计：\n"
                f"- 原始数据：{initial_count} 条\n"
                f"- 过滤无效数据：{filtered_count} 条（M列状态为已结束/utg）\n"
                f"- 有效数据：{len(filtered_df)} 条"
            )
        else:
            messagebox.showinfo("导入成功", f"毛坯基础信息导入成功！共{len(filtered_df)}条有效数据（M列无无效状态记录）")
        
        status_var.set(f"毛坯基础信息导入成功 - 有效数据{len(filtered_df)}条（过滤{filtered_count}条）")
        
    except Exception as e:
        messagebox.showerror("错误", f"毛坯基础信息导入失败：{str(e)}")
        status_var.set(f"毛坯基础信息导入失败：{str(e)}")

# ========== 新增：拆分后的每周订单导入函数 ==========
def get_token():
    """
    获取当前token，如果不存在则返回默认值
    """
    token_file = './token.json'
    if os.path.exists(token_file):
        try:
            with open(token_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data.get('token', 'ba777f9b9a0940cbaa1a6abcc9cb8ac5')
        except Exception:
            return 'ba777f9b9a0940cbaa1a6abcc9cb8ac5'
    return 'ba777f9b9a0940cbaa1a6abcc9cb8ac5'

def save_token(token):
    """
    保存token到文件
    """
    token_file = './token.json'
    with open(token_file, 'w', encoding='utf-8') as f:
        json.dump({'token': token}, f)

def ask_for_token():
    """
    弹出对话框让用户输入新的token
    返回用户输入的token，如果用户取消则返回None
    """
    from tkinter import simpledialog, Tk
    
    # 创建临时根窗口
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    
    # 弹出输入对话框
    token = simpledialog.askstring(
        "Token更新",
        "当前Token已过期，请输入新的Token:",
        parent=root
    )
    
    # 销毁临时窗口
    root.destroy()
    
    return token

def download_and_process_daily_order():
    """
    自动下载每日订单文件并处理数据
    返回处理后的DataFrame，如果失败返回None
    """
    # print("开始下载每日订单文件...")
    
    # 获取当前token
    current_token = get_token()
    # print(f"当前Token: {current_token[:10]}...")  # 只打印前10个字符，保护隐私
    
    # 尝试下载，如果token过期则更新
    max_attempts = 2  # 最多尝试2次（初始+1次更新token后重试）
    for attempt in range(max_attempts):
        # print(f"尝试下载，第{attempt+1}次...")
        
        try:
            # 构建请求头
            headers = {
                'accept': 'application/json, text/plain, */*',
                'accept-encoding': 'gzip, deflate',
                'accept-language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
                'connection': 'keep-alive',
                'cookie': f'token={current_token}',
                'host': '61.153.238.198:9331',
                'referer': 'http://61.153.238.198:9331/',
                'token': current_token,
                'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/143.0.0.0 Safari/537.36 Edg/143.0.0.0'
            }
            # print("请求头构建完成")
            
            # 生成当前时间戳
            import time
            timestamp = int(time.time() * 1000)
            # print(f"时间戳: {timestamp}")
            
            # 构建请求URL
            url = f'http://61.153.238.198:9331/api-zsdae/bu/jrgyspurorder/exportExcel'
            params = {
                't': timestamp,
                'page': 1,
                'limit': 2000,
                'key': '',
                'fentryStatus': 0,
                'supinfo': '',
                'khddh': '',
                'furl': 'http://61.153.238.198:9331/print-zsdae/api/export/excelExport'
            }
            # print(f"请求URL: {url}")
            # print(f"请求参数: {params}")
            
            # 发送请求
            # print("开始发送请求...")
            response = requests.get(url, headers=headers, params=params, verify=False, timeout=30)
            # print(f"请求完成，状态码: {response.status_code}")
            
            # 检查响应状态
            if response.status_code == 200:
                # print("响应状态码为200，开始处理数据...")
                
                # 检查响应内容大小
                content_length = len(response.content)
                # print(f"响应内容大小: {content_length} bytes")
                
                # 检查响应内容是否包含token失效信息
                try:
                    json_data = response.json()
                    # print("响应内容是JSON格式")
                    if isinstance(json_data, dict):
                        # 检查是否包含token失效信息
                        if json_data.get('code') == 401 or 'token失效' in json_data.get('msg', '') or '重新登录' in json_data.get('msg', ''):
                            # print("响应内容包含token失效信息")
                            
                            # 如果是第一次尝试，更新token并重试
                            if attempt < max_attempts - 1:
                                new_token = ask_for_token()
                                if new_token:
                                    # print("用户输入了新的Token")
                                    save_token(new_token)
                                    current_token = new_token
                                    continue  # 重试
                                else:
                                    # print("用户取消了Token更新")
                                    return None, None
                            else:
                                # print("Token更新后仍然失败")
                                return None, None
                except Exception:
                    # 响应内容不是JSON格式，继续处理
                    # print("响应内容不是JSON格式，继续处理...")
                
                # 如果内容太小，可能不是有效的Excel文件
                if content_length < 100:
                    # print("响应内容太小，可能不是有效的Excel文件")
                    # print(f"响应内容: {response.text}")
                    return None, None
                
                # 检查响应头中的内容类型
                content_type = response.headers.get('content-type', '')
                # print(f"响应内容类型: {content_type}")
                
                # 创建保存目录（使用绝对路径）
                save_dir = os.path.join(os.getcwd(), '每日订单自动下载')
                if not os.path.exists(save_dir):
                    try:
                        os.makedirs(save_dir, exist_ok=True)
                        # print(f"创建目录成功: {save_dir}")
                    except Exception as dir_e:
                        # print(f"创建目录失败: {str(dir_e)}")
                        # 目录创建失败，尝试直接处理响应内容
                
                # 生成文件名
                today = datetime.datetime.now().strftime('%Y%m%d')
                file_name = f'{today}_采购订单.xlsx'
                file_path = os.path.join(save_dir, file_name)
                
                # 尝试直接从响应内容创建DataFrame，避免文件保存权限问题
                # print("尝试直接从响应内容创建DataFrame...")
                try:
                    # 将响应内容保存到内存中，然后读取
                    from io import BytesIO
                    excel_data = BytesIO(response.content)
                    df = pd.read_excel(excel_data, engine='openpyxl')
                    # print(f"DataFrame创建成功，形状: {df.shape}")
                    
                    # 尝试保存文件（可选，即使失败也不影响数据处理）
                    try:
                        with open(file_path, 'wb') as f:
                            f.write(response.content)
                        # print(f"文件保存成功: {file_path}")
                    except Exception as save_e:
                        # print(f"文件保存失败（不影响数据处理）: {str(save_e)}")
                        # 保存失败但数据处理成功，继续执行
                    
                    return df, file_path
                except Exception as read_e:
                    # print(f"直接读取响应内容失败: {str(read_e)}")
                    import traceback
                    traceback.print_exc()
                    
                    # 尝试使用xlrd引擎
                    try:
                        from io import BytesIO
                        excel_data = BytesIO(response.content)
                        df = pd.read_excel(excel_data, engine='xlrd')
                        # print(f"DataFrame创建成功，形状: {df.shape}")
                        
                        # 尝试保存文件
                        try:
                            with open(file_path, 'wb') as f:
                                f.write(response.content)
                            # print(f"文件保存成功: {file_path}")
                        except Exception as save_e:
                            # print(f"文件保存失败（不影响数据处理）: {str(save_e)}")
                        
                        return df, file_path
                    except Exception as e:
                        # print(f"所有读取尝试都失败: {str(e)}")
                        return None, None
            elif response.status_code == 401:
                # print("响应状态码为401，Token已过期")
                
                # 如果是第一次尝试，更新token并重试
                if attempt < max_attempts - 1:
                    new_token = ask_for_token()
                    if new_token:
                        # print("用户输入了新的Token")
                        save_token(new_token)
                        current_token = new_token
                        continue  # 重试
                    else:
                        # print("用户取消了Token更新")
                        return None, None
                else:
                    # print("Token更新后仍然失败")
                    return None, None
            else:
                # print(f"下载失败，状态码: {response.status_code}")
                # print(f"响应内容: {response.text[:500]}...")
                return None, None
        except Exception as e:
            # print(f"下载过程中发生错误: {str(e)}")
            import traceback
            traceback.print_exc()
            return None, None
    
    return None, None

def import_daily_order_manual(global_data, status_var):
    """自动下载并导入每日订单文件"""
    try:
        # 下载并处理每日订单文件
        status_var.set("正在下载每日订单文件...")
        daily_order_df, daily_order_path = download_and_process_daily_order()
        
        if daily_order_df is None:
            messagebox.showerror("错误", "每日订单文件下载或处理失败！")
            status_var.set("每日订单导入失败 - 下载或处理失败")
            return
        
        # 保存上次导入的路径
        if daily_order_path:
            global_data.setdefault("last_import_paths", {})["daily_order"] = daily_order_path
        
        # 直接使用处理后的数据
        global_data["daily_order_df"] = daily_order_df
        
        update_overview_table(global_data["tree_overview"], "每日订单", len(global_data["daily_order_df"]))
        
        messagebox.showinfo("成功", "每日订单自动下载并导入成功！")
        status_var.set(f"每日订单导入成功 - 共{len(global_data['daily_order_df'])}条数据")
    except Exception as e:
        messagebox.showerror("错误", f"每日订单导入失败：{str(e)}")
        status_var.set(f"每日订单导入失败：{str(e)}")

# 恢复原来的每周订单手动导入函数
def import_weekly_order_manual(global_data, status_var):
    """手动选择每周订单文件"""
    weekly_order_path = filedialog.askopenfilename(
        title="选择每周订单表",
        filetypes=( ("Excel文件", "*.xlsx"), ("所有文件", "*.*")))
    if not weekly_order_path:
        return
    
    try:
        # 保存上次导入的路径
        global_data.setdefault("last_import_paths", {})["weekly_order"] = weekly_order_path
        
        # 只读取"喷涂订单整理"工作表
        try:
            # 检查文件中的工作表
            excel_file = pd.ExcelFile(weekly_order_path)
            if "喷涂订单整理" in excel_file.sheet_names:
                global_data["weekly_order_df"] = pd.read_excel(weekly_order_path, sheet_name="喷涂订单整理")
            else:
                messagebox.showerror("错误", "每周订单表中没有找到'喷涂订单整理'工作表！")
                status_var.set("每周订单导入失败 - 找不到指定工作表")
                return
        except Exception as inner_e:
            messagebox.showerror("错误", f"读取工作表失败：{str(inner_e)}")
            status_var.set(f"每周订单导入失败 - 读取工作表错误")
            return
        
        update_overview_table(global_data["tree_overview"], "每周订单", len(global_data["weekly_order_df"]))
        
        messagebox.showinfo("成功", "每周订单导入成功！")
        status_var.set(f"每周订单导入成功 - 共{len(global_data['weekly_order_df'])}条数据")
    except Exception as e:
        messagebox.showerror("错误", f"每周订单导入失败：{str(e)}")
        status_var.set(f"每周订单导入失败：{str(e)}")

# ========== 新增：当日成品库存导入函数 ==========
# ========== 当日成品库存手动导入函数（汇总+排除合计行） ==========
def import_finished_stock_manual(global_data, status_var):
    """
    手动选择当日成品库存文件，处理逻辑：
    1. 读取Excel文件，排除最后一行合计行
    2. 按A列（物料编码）分组，汇总G列（库存数量）求和
    3. 去重合并为单条记录，统计汇总后的物料编码数量
    """
    finished_stock_path = filedialog.askopenfilename(
        title="选择当日成品库存表",
        filetypes=( ("Excel文件", "*.xlsx"), ("所有文件", "*.*")))
    if not finished_stock_path:
        return
    
    try:
        # 保存上次导入的路径
        global_data.setdefault("last_import_paths", {})["finished_stock"] = finished_stock_path
        
        # 1. 读取成品库存表格（默认第一个工作表）
        stock_df = pd.read_excel(finished_stock_path)
        # 过滤全空行（先清理无效空行）
        stock_df = stock_df.dropna(how="all")
        
        # ========== 新增：排除最后一行合计行 ==========
        if len(stock_df) > 0:
            # 方案1：直接删除最后一行（通用，无论合计行内容）
            stock_df = stock_df.iloc[:-1, :]  # 取除最后一行外的所有行
            
            # 方案2：智能识别合计行（可选，防止最后一行不是合计的情况）
            # last_row = stock_df.iloc[-1]
            # # 判断最后一行A列是否包含"合计"/"总计"等关键词
            # if str(last_row.iloc[0]).strip() in ["合计", "总计", "合计行"] or "合计" in str(last_row.iloc[0]):
            #     stock_df = stock_df.iloc[:-1, :]
            # else:
            #     # 最后一行不是合计，保留
            #     pass
        
        # 2. 校验必要列是否存在（A列=物料编码，G列=库存数量）
        if stock_df.shape[1] < 7:
            messagebox.showerror("错误", "成品库存表格列数不足，缺少A列/ G列！")
            status_var.set("成品库存导入失败：列数不足")
            return
        
        # 3. 定义列名（按位置取列，避免列名变动）
        # A列=物料编码（第1列→索引0），G列=库存数量（第7列→索引6）
        stock_df.rename(columns={
            stock_df.columns[0]: "物料编码",
            stock_df.columns[6]: "库存数量"
        }, inplace=True)
        
        # 4. 数据清洗：过滤物料编码为空的行，库存数量转为数值型
        # 过滤物料编码空行
        stock_df = stock_df[stock_df["物料编码"].notna()]
        # 库存数量转为数值型（处理非数字的情况）
        stock_df["库存数量"] = pd.to_numeric(stock_df["库存数量"], errors="coerce").fillna(0)
        
        # 5. 按物料编码分组求和
        # 保留物料编码列，对库存数量求和
        summary_df = stock_df.groupby("物料编码", as_index=False).agg({
            "库存数量": "sum"  # 库存数量求和
            # 若需保留其他列（如物料名称），添加："物料名称": "first"
        })
        
        # 6. 统计信息
        original_rows = len(stock_df)          # 原始有效行数（排除合计+去空后）
        unique_material_count = len(summary_df)# 汇总后物料编码数量
        total_stock = summary_df["库存数量"].sum() # 总库存数量
        
        # 7. 存储汇总后的数据到全局变量
        global_data["finished_stock_df"] = summary_df
        
        # 8. 更新概览表格（显示汇总后的物料编码数量）
        update_overview_table(global_data["tree_overview"], "当日成品库存", unique_material_count)
        
        # 9. 提示成功信息（新增排除合计行说明）
        success_msg = (
            f"成品库存导入并汇总成功！\n"
            f"已排除最后一行合计行\n"
            f"原始有效行数：{original_rows} 行\n"
            f"汇总后物料编码数量：{unique_material_count} 个\n"
            f"总库存数量：{total_stock:.0f} 件"
        )
        messagebox.showinfo("成功", success_msg)
        status_var.set(f"成品库存导入成功 - 汇总后{unique_material_count}个物料编码（已排除合计行）")
    
    except Exception as e:
        messagebox.showerror("错误", f"成品库存导入失败：{str(e)}")
        status_var.set(f"成品库存导入失败：{str(e)}")
# ========== 新增：当日毛坯库存导入函数 ==========
# ========== 当日毛坯库存手动导入函数（汇总+排除合计行） ==========
def import_raw_stock_manual(global_data, status_var):
    """
    手动选择当日毛坯库存文件，处理逻辑：
    1. 读取Excel文件，排除最后一行合计行
    2. 按A列（物料编码）分组，汇总G列（库存数量）求和
    3. 去重合并为单条记录，统计汇总后的物料编码数量
    """
    raw_stock_path = filedialog.askopenfilename(
        title="选择当日毛坯库存表",
        filetypes=( ("Excel文件", "*.xlsx"), ("所有文件", "*.*")))
    if not raw_stock_path:
        return
    
    try:
        # 保存上次导入的路径
        global_data.setdefault("last_import_paths", {})["raw_stock"] = raw_stock_path
        
        # 1. 读取毛坯库存表格（默认第一个工作表）
        raw_df = pd.read_excel(raw_stock_path)
        # 过滤全空行（先清理无效空行）
        raw_df = raw_df.dropna(how="all")
        
        # ========== 排除最后一行合计行 ==========
        if len(raw_df) > 0:
            # 方案1：直接删除最后一行（通用，适配所有合计行在末尾的场景）
            raw_df = raw_df.iloc[:-1, :]
            
            # 方案2：智能识别合计行（可选，防止最后一行非合计行）
            # last_row = raw_df.iloc[-1]
            # if str(last_row.iloc[0]).strip() in ["合计", "总计"] or "合计" in str(last_row.iloc[0]):
            #     raw_df = raw_df.iloc[:-1, :]
        
        # 2. 校验列数是否足够（至少7列，保证A/G列存在）
        if raw_df.shape[1] < 7:
            messagebox.showerror("错误", "毛坯库存表格列数不足，缺少A列/G列！")
            status_var.set("毛坯库存导入失败：列数不足")
            return
        
        # 3. 列重命名（按位置，避免列名变动）
        # A列=物料编码（索引0），G列=库存数量（索引6）
        raw_df.rename(columns={
            raw_df.columns[0]: "物料编码",
            raw_df.columns[6]: "库存数量"
        }, inplace=True)
        
        # 4. 数据清洗
        raw_df = raw_df[raw_df["物料编码"].notna()]  # 过滤空编码行
        # 库存数量转为数值型，非数字值填充0
        raw_df["库存数量"] = pd.to_numeric(raw_df["库存数量"], errors="coerce").fillna(0)
        
        # 5. 按物料编码分组求和
        summary_df = raw_df.groupby("物料编码", as_index=False).agg({
            "库存数量": "sum"
            # 如需保留物料名称等列，添加："物料名称": "first"
        })
        
        # 6. 统计信息
        original_rows = len(raw_df)
        unique_material_count = len(summary_df)
        total_stock = summary_df["库存数量"].sum()
        
        # 7. 存储汇总后的数据到全局变量
        global_data["raw_stock_df"] = summary_df
        
        # 8. 更新概览表格和状态提示
        update_overview_table(global_data["tree_overview"], "当日毛坯库存", unique_material_count)
        
        success_msg = (
            f"毛坯库存导入并汇总成功！\n"
            f"已排除最后一行合计行\n"
            f"原始有效行数：{original_rows} 行\n"
            f"汇总后物料编码数量：{unique_material_count} 个\n"
            f"总毛坯库存数量：{total_stock:.0f} 件"
        )
        messagebox.showinfo("成功", success_msg)
        status_var.set(f"毛坯库存导入成功 - 汇总后{unique_material_count}个物料编码（已排除合计行）")
    
    except Exception as e:
        messagebox.showerror("错误", f"毛坯库存导入失败：{str(e)}")
        status_var.set(f"毛坯库存导入失败：{str(e)}")

# ========== 新增：年预计量手动导入函数 ==========
def import_yearly_estimate_manual(global_data, status_var):
    """手动选择年预计量文件"""
    yearly_estimate_path = filedialog.askopenfilename(
        title="选择年预计量表",
        filetypes=( ("Excel文件", "*.xlsx"), ("所有文件", "*.*")))
    if not yearly_estimate_path:
        return
    
    try:
        # 保存上次导入的路径
        global_data.setdefault("last_import_paths", {})["yearly_estimate"] = yearly_estimate_path
        
        # 读取年预计量数据并存储到全局变量
        
        # 获取Excel文件的所有工作表
        excel_file = pd.ExcelFile(yearly_estimate_path)
        sheet_names = excel_file.sheet_names

        
        # 选择工作表
        if len(sheet_names) == 1:
            # 只有一个工作表，直接使用
            selected_sheet = sheet_names[0]

        else:
            # 创建工作表选择对话框
            sheet_dialog = tk.Toplevel()
            sheet_dialog.title("选择工作表")
            sheet_dialog.geometry("400x300")
            sheet_dialog.resizable(False, False)
            
            # 居中显示
            sheet_dialog.update_idletasks()
            width = sheet_dialog.winfo_width()
            height = sheet_dialog.winfo_height()
            x = (sheet_dialog.winfo_screenwidth() // 2) - (width // 2)
            y = (sheet_dialog.winfo_screenheight() // 2) - (height // 2)
            sheet_dialog.geometry('{}x{}+{}+{}'.format(width, height, x, y))
            
            # 添加标题
            ttk.Label(sheet_dialog, text="请选择要使用的工作表：", font=("微软雅黑", 12, "bold")).pack(pady=10)
            
            # 添加列表框
            listbox_frame = ttk.Frame(sheet_dialog)
            listbox_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
            
            sheet_listbox = tk.Listbox(listbox_frame, font=("微软雅黑", 10))
            scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=sheet_listbox.yview)
            sheet_listbox.configure(yscrollcommand=scrollbar.set)
            
            sheet_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            # 添加工作表名称到列表框
            for sheet in sheet_names:
                sheet_listbox.insert(tk.END, sheet)
            
            # 默认选择第一个工作表
            sheet_listbox.select_set(0)
            sheet_listbox.focus_set()
            
            # 选择结果变量
            selected_sheet = [sheet_names[0]]  # 使用列表存储，以便在回调中修改
            
            # 确认按钮回调
            def on_confirm():
                if sheet_listbox.curselection():
                    selected_sheet[0] = sheet_names[sheet_listbox.curselection()[0]]
                    sheet_dialog.destroy()
            
            # 取消按钮回调
            def on_cancel():
                selected_sheet[0] = None
                sheet_dialog.destroy()
            
            # 添加按钮
            button_frame = ttk.Frame(sheet_dialog)
            button_frame.pack(pady=10)
            
            ttk.Button(button_frame, text="确认", command=on_confirm, width=10).pack(side=tk.LEFT, padx=5)
            ttk.Button(button_frame, text="取消", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)
            
            # 等待对话框关闭
            sheet_dialog.wait_window()
            
            # 如果用户取消，返回
            if selected_sheet[0] is None:
    
                return
            

        
        # 读取第一个工作表
        
        # 尝试不同的读取方式
        try:
            # 方式1：默认方式读取

            df_default = pd.read_excel(excel_file, sheet_name=selected_sheet[0])

            if len(df_default) > 0:
                pass
        except Exception as e:
            pass
        
        try:
            # 方式2：指定dtype为str

            df_str = pd.read_excel(excel_file, sheet_name=selected_sheet[0], dtype=str)

            if len(df_str) > 0:
                pass
        except Exception as e:
            pass
        
        try:
            # 方式3：指定engine为openpyxl

            df_openpyxl = pd.read_excel(excel_file, sheet_name=selected_sheet[0], engine='openpyxl', dtype=str)

            if len(df_openpyxl) > 0:
                pass
        except Exception as e:
            pass
        
        # 使用方式2的结果继续处理
        global_data["yearly_estimate_df"] = df_str
        yearly_df = global_data["yearly_estimate_df"]

        

        
        
        
        
        
        # 尝试筛选特定成品编码
        target_code = "1.2.1.0103080010"
        
        
        
        
        
        # 检查是否有匹配的行
        filtered = yearly_df[yearly_df.iloc[:, 2].astype(str).str.strip() == target_code]
        if not filtered.empty:
            # 计算G列的数值总和
            numeric_values = pd.to_numeric(filtered.iloc[:, 6], errors="coerce")
        else:
            # 未找到精确匹配，尝试模糊查找
            fuzzy_filtered = yearly_df[yearly_df.iloc[:, 2].str.contains(target_code, na=False)]
            if not fuzzy_filtered.empty:
                pass
            else:
                pass


        
        # 更新概览表格
        update_overview_table(global_data["tree_overview"], "年预计量", len(yearly_df))
        
        # 提示成功信息
        if len(sheet_names) == 1:
            messagebox.showinfo("成功", f"年预计量导入成功！\n已使用工作表：{selected_sheet}")
            status_var.set(f"年预计量导入成功 - 共{len(yearly_df)}条数据")
        else:
            messagebox.showinfo("成功", f"年预计量导入成功！\n已使用工作表：{selected_sheet[0]}")
            status_var.set(f"年预计量导入成功 - 共{len(yearly_df)}条数据 (工作表: {selected_sheet[0]})")
        

    except Exception as e:
        messagebox.showerror("错误", f"年预计量导入失败：{str(e)}")
        status_var.set(f"年预计量导入失败：{str(e)}")

# ========== 新增：已下单文件手动导入函数 ==========
# ========== 新增：快速导入所有上次文件的函数 ==========
def quick_import_all_last_files(global_data, status_var):
    """
    一次性导入所有有历史记录的文件
    """
    # 检查是否有上次导入的路径
    last_import_paths = global_data.get("last_import_paths", {})
    
    # 调试信息：显示所有历史路径
    # print("历史导入路径:", last_import_paths)
    # print("是否包含ordered_file:", "ordered_file" in last_import_paths)
    
    if not last_import_paths:
        messagebox.showwarning("提示", "暂无任何历史导入记录，请先手动选择文件导入一次！")
        return
    
    # 定义数据类型的导入顺序和对应的数据类型名称
    data_types_order = [
        ("raw_material", "毛坯基础信息"),
        ("weekly_order", "每周订单"),
        ("daily_order", "每日订单"),
        ("yearly_estimate", "年预计量"),
        ("ordered_file", "已下单文件"),
        ("finished_stock", "当日成品库存"),
        ("raw_stock", "当日毛坯库存")
    ]
    
    imported_count = 0
    failed_count = 0
    failed_types = []
    
    try:
        for data_type, display_name in data_types_order:
            if data_type in last_import_paths:
                path = last_import_paths[data_type]
                status_var.set(f"正在快速导入：{display_name}...")
                
                try:
                    if data_type == "raw_material":
                        # 复制毛坯基础信息导入的核心逻辑
                        raw_df = pd.read_excel(path)
                        initial_count = len(raw_df)
                        
                        if initial_count == 0:
                            messagebox.showwarning("警告", f"{display_name}表为空！")
                            status_var.set(f"{display_name}导入失败 - 表格为空")
                            failed_count += 1
                            failed_types.append(display_name)
                            continue
                        
                        # 核心：固定M列（索引12）过滤无效状态
                        if len(raw_df.columns) <= 12:
                            messagebox.showerror("错误", f"{display_name}表列数不足！M列（第13列）不存在")
                            status_var.set(f"{display_name}导入失败 - 无M列")
                            failed_count += 1
                            failed_types.append(display_name)
                            continue
                        
                        status_col_index = 12  # M列 = 索引12
                        status_col_name = raw_df.columns[status_col_index]
                        raw_df[status_col_name] = raw_df[status_col_name].astype(str).str.strip().str.lower()
                        
                        valid_mask = ~(
                            raw_df[status_col_name].str.contains("已结束", na=False) |
                            raw_df[status_col_name].str.contains("utg", na=False)
                        )
                        filtered_df = raw_df[valid_mask].copy()
                        filtered_count = initial_count - len(filtered_df)
                        
                        global_data["raw_material_df"] = filtered_df
                        update_overview_table(global_data["tree_overview"], "毛坯基础信息", len(filtered_df))
                        imported_count += 1
                        
                    elif data_type == "weekly_order":
                        # 复制每周订单导入的核心逻辑，只读取"喷涂订单整理"工作表
                        try:
                            # 检查文件中的工作表
                            excel_file = pd.ExcelFile(path)
                            if "喷涂订单整理" in excel_file.sheet_names:
                                global_data["weekly_order_df"] = pd.read_excel(path, sheet_name="喷涂订单整理")
                                update_overview_table(global_data["tree_overview"], "每周订单", len(global_data["weekly_order_df"]))
                                imported_count += 1
                            else:
                                messagebox.showerror("错误", f"{display_name}表中没有找到'喷涂订单整理'工作表！")
                                status_var.set(f"{display_name}导入失败 - 找不到指定工作表")
                                failed_count += 1
                                failed_types.append(display_name)
                                continue
                        except Exception as inner_e:
                            messagebox.showerror("错误", f"读取工作表失败：{str(inner_e)}")
                            status_var.set(f"{display_name}导入失败 - 读取工作表错误")
                            failed_count += 1
                            failed_types.append(display_name)
                            continue
                        
                    elif data_type == "daily_order":
                        # 自动下载并处理每日订单文件，与手动导入逻辑一致
                        try:
                            daily_order_df, daily_order_path = download_and_process_daily_order()
                            
                            if daily_order_df is None:
                                messagebox.showerror("错误", "每日订单文件下载或处理失败！")
                                status_var.set("每日订单导入失败 - 下载或处理失败")
                                failed_count += 1
                                failed_types.append(display_name)
                                continue
                            
                            # 保存上次导入的路径
                            if daily_order_path:
                                global_data.setdefault("last_import_paths", {})["daily_order"] = daily_order_path
                            
                            # 直接使用处理后的数据
                            global_data["daily_order_df"] = daily_order_df
                            update_overview_table(global_data["tree_overview"], "每日订单", len(global_data["daily_order_df"]))
                            imported_count += 1
                        except Exception as inner_e:
                            messagebox.showerror("错误", f"每日订单导入失败：{str(inner_e)}")
                            status_var.set(f"每日订单导入失败：{str(inner_e)}")
                            failed_count += 1
                            failed_types.append(display_name)
                            continue
                        
                    elif data_type == "finished_stock":
                        # 复制成品库存导入的核心逻辑
                        stock_df = pd.read_excel(path)
                        stock_df = stock_df.dropna(how="all")
                        
                        if len(stock_df) > 0:
                            stock_df = stock_df.iloc[:-1, :]
                        
                        if stock_df.shape[1] < 7:
                            messagebox.showerror("错误", f"{display_name}表格列数不足，缺少A列/ G列！")
                            status_var.set(f"{display_name}导入失败：列数不足")
                            failed_count += 1
                            failed_types.append(display_name)
                            continue
                        
                        stock_df.rename(columns={
                            stock_df.columns[0]: "物料编码",
                            stock_df.columns[6]: "库存数量"
                        }, inplace=True)
                        
                        stock_df = stock_df[stock_df["物料编码"].notna()]
                        stock_df["库存数量"] = pd.to_numeric(stock_df["库存数量"], errors="coerce").fillna(0)
                        
                        summary_df = stock_df.groupby("物料编码", as_index=False).agg({
                            "库存数量": "sum"
                        })
                        
                        global_data["finished_stock_df"] = summary_df
                        update_overview_table(global_data["tree_overview"], "当日成品库存", len(summary_df))
                        imported_count += 1
                        
                    elif data_type == "raw_stock":
                        # 复制毛坯库存导入的核心逻辑
                        raw_df = pd.read_excel(path)
                        raw_df = raw_df.dropna(how="all")
                        
                        if len(raw_df) > 0:
                            raw_df = raw_df.iloc[:-1, :]
                        
                        if raw_df.shape[1] < 7:
                            messagebox.showerror("错误", f"{display_name}表格列数不足，缺少A列/G列！")
                            status_var.set(f"{display_name}导入失败：列数不足")
                            failed_count += 1
                            failed_types.append(display_name)
                            continue
                        
                        raw_df.rename(columns={
                            raw_df.columns[0]: "物料编码",
                            raw_df.columns[6]: "库存数量"
                        }, inplace=True)
                        
                        raw_df = raw_df[raw_df["物料编码"].notna()]
                        raw_df["库存数量"] = pd.to_numeric(raw_df["库存数量"], errors="coerce").fillna(0)
                        
                        summary_df = raw_df.groupby("物料编码", as_index=False).agg({
                            "库存数量": "sum"
                        })
                        
                        global_data["raw_stock_df"] = summary_df
                        update_overview_table(global_data["tree_overview"], "当日毛坯库存", len(summary_df))
                        imported_count += 1
                        
                    elif data_type == "yearly_estimate":
                        # 复制年预计量导入的核心逻辑
                        excel_file = pd.ExcelFile(path)
                        sheet_names = excel_file.sheet_names
                        
                        if len(sheet_names) == 1:
                            selected_sheet = sheet_names[0]
                        else:
                            messagebox.showwarning("提示", f"{display_name}文件包含多个工作表，无法快速导入。请手动选择文件并选择工作表！")
                            failed_count += 1
                            failed_types.append(display_name)
                            continue
                        
                        try:
                            df_str = pd.read_excel(excel_file, sheet_name=selected_sheet, dtype=str)
                            global_data["yearly_estimate_df"] = df_str
                            imported_count += 1
                        except Exception as e:
                            messagebox.showerror("错误", f"{display_name}快速导入失败：{str(e)}")
                            failed_count += 1
                            failed_types.append(display_name)
                            continue
                        
                    elif data_type == "ordered_file":
                        # print("开始处理已下单文件:", path)
                        # 复制已下单文件导入的核心逻辑
                        # 确保能够读取xlsm格式文件
                        try:
                            excel_file = pd.ExcelFile(path, engine='openpyxl')
                            sheet_names = excel_file.sheet_names
                            # print("已下单文件包含工作表:", sheet_names)
                            
                            all_valid_df = []
                            total_valid_rows = 0
                            total_skipped_rows = 0
                            
                            for sheet in sheet_names:
                                # print("处理工作表:", sheet)
                                sheet_df = pd.read_excel(path, sheet_name=sheet, engine='openpyxl')
                                # print("工作表原始行数:", len(sheet_df))
                                sheet_df = sheet_df.dropna(how="all")
                                # print("过滤全空行后行数:", len(sheet_df))
                                
                                if len(sheet_df) == 0:
                                    # print("工作表没有有效行，跳过")
                                    continue
                                
                                # 检查列数是否足够
                                if len(sheet_df.columns) <= 6:
                                    # print(f"工作表列数不足，只有{len(sheet_df.columns)}列，需要至少7列")
                                    continue
                                
                                f_col = sheet_df.iloc[:, 5]  # F列
                                g_col = sheet_df.iloc[:, 6]  # G列
                                filter_condition = ~(f_col.notna() & g_col.notna())
                                
                                valid_rows = sheet_df[filter_condition]
                                skipped_rows = sheet_df[~filter_condition]
                                
                                # print(f"工作表{sheet}有效行数:", len(valid_rows))
                                # print(f"工作表{sheet}跳过行数:", len(skipped_rows))
                                
                                # 标记所属工作表，便于溯源
                                valid_rows["所属工作表"] = sheet
                                all_valid_df.append(valid_rows)
                                total_valid_rows += len(valid_rows)
                                total_skipped_rows += len(skipped_rows)
                            
                            # print("所有工作表处理完成")
                            # print("有效行总数:", total_valid_rows)
                            # print("跳过行总数:", total_skipped_rows)
                            
                            if all_valid_df:
                                merged_df = pd.concat(all_valid_df, ignore_index=True)
                                global_data["ordered_file_df"] = merged_df
                                imported_count += 1
                                # print("已下单文件导入成功")
                                # 更新概览表格
                                update_overview_table(global_data["tree_overview"], "已下单文件", total_valid_rows)
                            else:
                                # print("没有有效数据，导入失败")
                                failed_count += 1
                                failed_types.append(display_name)
                                continue
                        except Exception as e:
                            # print("处理已下单文件时出错:", str(e))
                            import traceback
                            traceback.print_exc()
                            messagebox.showerror("错误", f"{display_name}快速导入失败：{str(e)}")
                            failed_count += 1
                            failed_types.append(display_name)
                            continue
                except Exception as e:
                    messagebox.showerror("错误", f"{display_name}快速导入失败：{str(e)}")
                    failed_count += 1
                    failed_types.append(display_name)
                    continue
        
        # 显示导入结果
        success_msg = f"快速导入完成！\n"
        success_msg += f"成功导入：{imported_count} 个文件\n"
        if failed_count > 0:
            success_msg += f"导入失败：{failed_count} 个文件\n"
            success_msg += f"失败的文件类型：{', '.join(failed_types)}\n"
        
        messagebox.showinfo("导入结果", success_msg)
        status_var.set(f"快速导入完成 - 成功{imported_count}个，失败{failed_count}个")
        
    except Exception as e:
        messagebox.showerror("错误", f"快速导入过程中发生错误：{str(e)}")
        status_var.set(f"快速导入失败：{str(e)}")

# ========== 新增：快速导入上次文件的函数 ==========
def quick_import_from_last_path(global_data, status_var, data_type, import_func):
    """
    从上次导入的路径快速导入文件
    data_type: 数据类型名称（用于在last_import_paths中查找路径）
    import_func: 对应的完整导入函数
    """
    # 检查是否有上次导入的路径
    last_import_paths = global_data.get("last_import_paths", {})
    if data_type not in last_import_paths:
        messagebox.showwarning("提示", f"暂无{data_type}的历史导入记录，请先手动选择文件导入一次！")
        return
    
    try:
        # 直接调用对应的导入函数，模拟手动选择文件的流程
        # 注意：这里需要重新实现导入逻辑，因为原导入函数都包含文件选择对话框
        path = last_import_paths[data_type]
        
        if data_type == "raw_material":
            # 复制毛坯基础信息导入的核心逻辑
            raw_df = pd.read_excel(path)
            initial_count = len(raw_df)
            
            if initial_count == 0:
                messagebox.showwarning("警告", "毛坯基础信息表为空！")
                status_var.set("毛坯基础信息导入失败 - 表格为空")
                return
            
            # 核心：固定M列（索引12）过滤无效状态
            if len(raw_df.columns) <= 12:
                messagebox.showerror("错误", "毛坯基础信息表列数不足！M列（第13列）不存在")
                status_var.set("毛坯基础信息导入失败 - 无M列")
                return
            
            status_col_index = 12  # M列 = 索引12
            status_col_name = raw_df.columns[status_col_index]
            raw_df[status_col_name] = raw_df[status_col_name].astype(str).str.strip().str.lower()
            
            valid_mask = ~(
                raw_df[status_col_name].str.contains("已结束", na=False) |
                raw_df[status_col_name].str.contains("utg", na=False)
            )
            filtered_df = raw_df[valid_mask].copy()
            filtered_count = initial_count - len(filtered_df)
            
            global_data["raw_material_df"] = filtered_df
            update_overview_table(global_data["tree_overview"], "毛坯基础信息", len(filtered_df))
            
            if filtered_count > 0:
                messagebox.showinfo("导入成功", 
                    f"毛坯基础信息导入成功！\n"\
                    f"📊 数据统计：\n"\
                    f"- 原始数据：{initial_count} 条\n"\
                    f"- 过滤无效数据：{filtered_count} 条（M列状态为已结束/utg）\n"\
                    f"- 有效数据：{len(filtered_df)} 条"
                )
            else:
                messagebox.showinfo("导入成功", f"毛坯基础信息导入成功！共{len(filtered_df)}条有效数据（M列无无效状态记录）")
            
            status_var.set(f"毛坯基础信息导入成功 - 有效数据{len(filtered_df)}条（过滤{filtered_count}条）")
        
        elif data_type == "weekly_order":
            # 复制每周订单导入的核心逻辑，只读取"喷涂订单整理"工作表
            try:
                # 检查文件中的工作表
                excel_file = pd.ExcelFile(path)
                if "喷涂订单整理" in excel_file.sheet_names:
                    global_data["weekly_order_df"] = pd.read_excel(path, sheet_name="喷涂订单整理")
                    update_overview_table(global_data["tree_overview"], "每周订单", len(global_data["weekly_order_df"]))
                    
                    messagebox.showinfo("成功", "每周订单快速导入成功！")
                    status_var.set(f"每周订单快速导入成功 - 共{len(global_data['weekly_order_df'])}条数据")
                else:
                    messagebox.showerror("错误", "每周订单表中没有找到'喷涂订单整理'工作表！")
                    status_var.set("每周订单导入失败 - 找不到指定工作表")
                    return
            except Exception as inner_e:
                messagebox.showerror("错误", f"读取工作表失败：{str(inner_e)}")
                status_var.set(f"每周订单导入失败 - 读取工作表错误")
                return
        
        elif data_type == "finished_stock":
            # 复制成品库存导入的核心逻辑
            stock_df = pd.read_excel(path)
            stock_df = stock_df.dropna(how="all")
            
            if len(stock_df) > 0:
                stock_df = stock_df.iloc[:-1, :]
            
            if stock_df.shape[1] < 7:
                messagebox.showerror("错误", "成品库存表格列数不足，缺少A列/ G列！")
                status_var.set("成品库存导入失败：列数不足")
                return
            
            stock_df.rename(columns={
                stock_df.columns[0]: "物料编码",
                stock_df.columns[6]: "库存数量"
            }, inplace=True)
            
            stock_df = stock_df[stock_df["物料编码"].notna()]
            stock_df["库存数量"] = pd.to_numeric(stock_df["库存数量"], errors="coerce").fillna(0)
            
            summary_df = stock_df.groupby("物料编码", as_index=False).agg({
                "库存数量": "sum"
            })
            
            original_rows = len(stock_df)
            unique_material_count = len(summary_df)
            total_stock = summary_df["库存数量"].sum()
            
            global_data["finished_stock_df"] = summary_df
            update_overview_table(global_data["tree_overview"], "当日成品库存", unique_material_count)
            
            success_msg = (f"成品库存快速导入成功！\n"\
                        f"原始数据行数：{original_rows} 行\n"\
                        f"汇总后物料编码数：{unique_material_count} 个\n"\
                        f"总库存数量：{total_stock} 件")
            messagebox.showinfo("成功", success_msg)
            status_var.set(f"成品库存快速导入成功 - 汇总后{unique_material_count}个物料编码（已排除合计行）")
        
        elif data_type == "raw_stock":
            # 复制毛坯库存导入的核心逻辑
            raw_df = pd.read_excel(path)
            raw_df = raw_df.dropna(how="all")
            
            if len(raw_df) > 0:
                raw_df = raw_df.iloc[:-1, :]
            
            if raw_df.shape[1] < 7:
                messagebox.showerror("错误", "毛坯库存表格列数不足，缺少A列/G列！")
                status_var.set("毛坯库存导入失败：列数不足")
                return
            
            raw_df.rename(columns={
                raw_df.columns[0]: "物料编码",
                raw_df.columns[6]: "库存数量"
            }, inplace=True)
            
            raw_df = raw_df[raw_df["物料编码"].notna()]
            raw_df["库存数量"] = pd.to_numeric(raw_df["库存数量"], errors="coerce").fillna(0)
            
            summary_df = raw_df.groupby("物料编码", as_index=False).agg({
                "库存数量": "sum"
            })
            
            original_rows = len(raw_df)
            unique_material_count = len(summary_df)
            total_stock = summary_df["库存数量"].sum()
            
            global_data["raw_stock_df"] = summary_df
            update_overview_table(global_data["tree_overview"], "当日毛坯库存", unique_material_count)
            
            success_msg = (f"毛坯库存快速导入成功！\n"\
                        f"原始数据行数：{original_rows} 行\n"\
                        f"汇总后物料编码数：{unique_material_count} 个\n"\
                        f"总库存数量：{total_stock} 件")
            messagebox.showinfo("成功", success_msg)
            status_var.set(f"毛坯库存快速导入成功 - 汇总后{unique_material_count}个物料编码（已排除合计行）")
        
        elif data_type == "yearly_estimate":
            # 复制年预计量导入的核心逻辑
            excel_file = pd.ExcelFile(path)
            sheet_names = excel_file.sheet_names
            
            if len(sheet_names) == 1:
                selected_sheet = sheet_names[0]
            else:
                messagebox.showwarning("提示", "年预计量文件包含多个工作表，无法快速导入。请手动选择文件并选择工作表！")
                return
            
            try:
                df_str = pd.read_excel(excel_file, sheet_name=selected_sheet, dtype=str)
                global_data["yearly_estimate_df"] = df_str
                yearly_df = global_data["yearly_estimate_df"]
                
                messagebox.showinfo("成功", f"年预计量快速导入成功！\n已使用工作表：{selected_sheet}")
                status_var.set(f"年预计量快速导入成功 - 共{len(yearly_df)}条数据 (工作表: {selected_sheet})")
            except Exception as e:
                messagebox.showerror("错误", f"年预计量快速导入失败：{str(e)}")
                status_var.set(f"年预计量快速导入失败：{str(e)}")
        
        elif data_type == "ordered_file":
            # 复制已下单文件导入的核心逻辑
            excel_file = pd.ExcelFile(path, engine='openpyxl')
            sheet_names = excel_file.sheet_names
            
            all_valid_df = []
            total_valid_rows = 0
            total_skipped_rows = 0
            
            for sheet in sheet_names:
                sheet_df = pd.read_excel(path, sheet_name=sheet, engine='openpyxl')
                sheet_df = sheet_df.dropna(how="all")
                
                f_col = sheet_df.iloc[:, 5]  # F列
                g_col = sheet_df.iloc[:, 6]  # G列
                filter_condition = ~(f_col.notna() & g_col.notna())
                
                valid_rows = sheet_df[filter_condition]
                skipped_rows = sheet_df[~filter_condition]
                
                all_valid_df.append(valid_rows)
                total_valid_rows += len(valid_rows)
                total_skipped_rows += len(skipped_rows)
            
            if all_valid_df:
                merged_df = pd.concat(all_valid_df, ignore_index=True)
                global_data["ordered_file_df"] = merged_df
            
            success_msg = (
                f"已下单文件快速导入成功！\n"\
                f"工作簿包含 {len(sheet_names)} 个工作表\n"\
                f"过滤后有效行数：{total_valid_rows} 行\n"\
                f"跳过F/G列都有值的行数：{total_skipped_rows} 行"
            )
            messagebox.showinfo("成功", success_msg)
            status_var.set(f"已下单文件快速导入成功 - 有效行数{total_valid_rows}（跳过{total_skipped_rows}行）")
            
    except Exception as e:
        messagebox.showerror("错误", f"快速导入失败：{str(e)}")
        status_var.set(f"快速导入失败：{str(e)}")

# ========== 已下单文件手动导入函数（过滤F/G列都有值的行） ==========
def import_ordered_file_manual(global_data, status_var):
    """
    手动选择已下单文件，统计规则：
    1. 遍历整个工作簿所有工作表
    2. 过滤：F列和G列都有值的行跳过，仅统计至少一列为空的行
    3. 合并所有符合条件的数据并统计总行数
    """
    ordered_file_path = filedialog.askopenfilename(
        title="选择已下单文件",
        filetypes=( ("Excel文件", "*.xlsm"), ("所有文件", "*.*")))
    if not ordered_file_path:
        return
    
    try:
        # 保存上次导入的路径
        global_data.setdefault("last_import_paths", {})["ordered_file"] = ordered_file_path
        
        # 1. 获取Excel工作簿的所有工作表名称
        excel_file = pd.ExcelFile(ordered_file_path, engine='openpyxl')
        sheet_names = excel_file.sheet_names
        
        # 打印调试信息
        # print("手动导入已下单文件 - 文件路径:", ordered_file_path)
        # print("手动导入已下单文件 - 工作表名称:", sheet_names)
        
        # 2. 遍历所有工作表，过滤并合并数据
        all_valid_df = []  # 存储所有符合条件的行
        total_valid_rows = 0  # 统计过滤后的总行数
        total_skipped_rows = 0  # 统计跳过的行数（F/G都有值）
        
        for sheet in sheet_names:
            # 读取当前工作表（保留所有行，后续过滤）
            sheet_df = pd.read_excel(ordered_file_path, sheet_name=sheet, engine='openpyxl')
            # print(f"手动导入已下单文件 - 工作表{sheet}原始行数:", len(sheet_df))
            # 过滤全空行（先清理无效空行）
            sheet_df = sheet_df.dropna(how="all")
            
            # 关键：过滤F列和G列都有值的行
            # 方案1：按列名过滤（若F/G列有明确列名，比如"列F"/"列G"，替换下面的列名）
            # filter_condition = ~(sheet_df["F列名"].notna() & sheet_df["G列名"].notna())
            
            # 方案2：按列位置过滤（F列=第6列，G列=第7列，索引从0开始则为5/6）
            # 优先推荐按位置（避免列名变动）
            f_col = sheet_df.iloc[:, 5]  # F列（第6列）
            g_col = sheet_df.iloc[:, 6]  # G列（第7列）
            # 过滤条件：F/G不同时有值（即至少一列为空）
            filter_condition = ~(f_col.notna() & g_col.notna())
            
            # 筛选符合条件的行
            valid_df = sheet_df[filter_condition]
            skipped_rows = len(sheet_df) - len(valid_df)
            
            # 累加统计数
            total_valid_rows += len(valid_df)
            total_skipped_rows += skipped_rows
            
            # 标记所属工作表，便于溯源
            valid_df["所属工作表"] = sheet
            all_valid_df.append(valid_df)
        
        # 3. 合并所有符合条件的数据（存入全局变量）
        if all_valid_df:
            global_data["ordered_file_df"] = pd.concat(all_valid_df, ignore_index=True)
        else:
            global_data["ordered_file_df"] = pd.DataFrame()
        
        # 4. 更新概览表格和状态提示
        update_overview_table(global_data["tree_overview"], "已下单文件", total_valid_rows)
        
        # 详细提示信息（包含过滤统计）
        success_msg = (
            f"已下单文件导入成功！\n"
            f"工作簿包含 {len(sheet_names)} 个工作表\n"
            f"过滤后有效行数：{total_valid_rows} 行\n"
            f"跳过F/G列都有值的行数：{total_skipped_rows} 行"
        )
        messagebox.showinfo("成功", success_msg)
        status_var.set(f"已下单文件导入成功 - 有效行数{total_valid_rows}（跳过{total_skipped_rows}行）")
        
        # 更新订单追踪表格
        from ui_modules import update_order_tracking_table
        update_order_tracking_table(global_data)
    
    except Exception as e:
        messagebox.showerror("错误", f"已下单文件导入失败：{str(e)}")
        status_var.set(f"已下单文件导入失败：{str(e)}")

def save_analysis_report(global_data, path_config):
    """保存分析报告"""
    if global_data["analysis_result_df"] is None:
        messagebox.showwarning("警告", "暂无分析结果可保存！")
        return
    
    # 选择保存路径
    save_path = filedialog.asksaveasfilename(
        title="保存分析报告",
        initialdir=path_config["report"],
        initialfile=f"采购分析_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
        filetypes=(("Excel文件", "*.xlsx"), ("所有文件", "*.*"))
    )
    if not save_path:
        return
    
    try:
        global_data["analysis_result_df"].to_excel(save_path, index=False)
        messagebox.showinfo("成功", f"报告已保存到：{save_path}")
        # 刷新历史报告列表（需重新加载）
    except Exception as e:
        messagebox.showerror("错误", f"报告保存失败：{str(e)}")

def load_history_reports(tree_history, report_path):
    """加载历史报告列表"""
    # 清空表格
    for item in tree_history.get_children():
        tree_history.delete(item)
    
    # 读取报告文件夹
    if os.path.exists(report_path):
        for file in os.listdir(report_path):
            if file.endswith(".xlsx") and file.startswith("采购分析_"):
                file_path = os.path.join(report_path, file)
                mtime = os.path.getmtime(file_path)
                mtime_str = pd.Timestamp.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
                tree_history.insert("", "end", values=(file, mtime_str, "打开"))

def update_overview_table(tree, data_type, count):
    """更新数据概览表格"""
    for item in tree.get_children():
        values = tree.item(item)["values"]
        if values[0] == data_type:
            tree.item(item, values=(
                data_type,
                f"{count}条",
                pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")
            ))
            break
