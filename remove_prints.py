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
            lines = f.readlines()
        
        # 处理每一行，移除 print 语句
        output_lines = []
        for line in lines:
            # 如果行包含 print 语句，将其注释掉
            if 'print(' in line:
                # 保留行首的缩进
                indent = ''
                for char in line:
                    if char in (' ', '\t'):
                        indent += char
                    else:
                        break
                # 注释掉 print 语句
                output_lines.append(f"{indent}# {line.lstrip()}")
            else:
                output_lines.append(line)
        
        # 保存处理后的文件
        with open(output_file, 'w', encoding='utf-8') as f:
            f.writelines(output_lines)
        
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
