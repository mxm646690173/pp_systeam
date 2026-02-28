#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
复制文件并覆盖已存在的文件
"""

import shutil
import os

def copy_file(source, destination):
    """
    复制文件并覆盖已存在的文件
    """
    try:
        # 复制文件
        shutil.copy2(source, destination)
        print(f"成功将文件从 {source} 复制到 {destination}")
        return True
    except Exception as e:
        print(f"复制文件失败: {e}")
        return False

if __name__ == "__main__":
    source = "data_handling_commented.py"
    destination = "data_handling.py"
    
    if os.path.exists(source):
        copy_file(source, destination)
    else:
        print(f"源文件不存在: {source}")
