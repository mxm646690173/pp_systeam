#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
移除 data_handling.py 文件中的所有 print 语句
"""

import os

def remove_prints(input_file, output_file):
    """
    移除文件中的所有 print 语句
    """
    # 尝试读取文件
    try:
        with open(input_file, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        
        # 替换所有的 print 语句
        import re
        # 使用正则表达式匹配 print 语句，保留缩进
        pattern = r'(^\s*)(print\([^)]*\))'
        content = re.sub(pattern, r'\1# \2', content, flags=re.MULTILINE)
        
        # 保存处理后的文件
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print(f"成功移除所有 print 语句，保存到: {output_file}")
        return True
    except Exception as e:
        print(f"处理文件失败: {e}")
        return False

if __name__ == "__main__":
    input_file = "data_handling.py"
    output_file = "data_handling_no_prints.py"
    
    if os.path.exists(input_file):
        remove_prints(input_file, output_file)
    else:
        print(f"文件不存在: {input_file}")
