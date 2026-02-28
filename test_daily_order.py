import sys
import os

# 添加当前目录到系统路径
sys.path.append(os.getcwd())

from data_handling import download_and_process_daily_order

print("测试每日订单导入功能...")
print("=" * 50)

# 调用函数
daily_order_df, daily_order_path = download_and_process_daily_order()

print("=" * 50)
if daily_order_df is not None:
    print(f"成功：每日订单导入成功！")
    print(f"文件路径：{daily_order_path}")
    print(f"数据形状：{daily_order_df.shape}")
else:
    print("失败：每日订单导入失败！")
