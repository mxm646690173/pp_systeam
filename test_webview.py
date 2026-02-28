import tkinter as tk

# 尝试导入WebView小部件
try:
    from tkinter import WebView
    print("tkinter有WebView小部件")
except ImportError:
    print("tkinter没有WebView小部件")

# 尝试导入ttk的WebView小部件
try:
    from tkinter.ttk import WebView
    print("ttk有WebView小部件")
except ImportError:
    print("ttk没有WebView小部件")

# 尝试导入其他可能的WebView模块
try:
    import webview
    print("webview模块可用")
except ImportError:
    print("webview模块不可用")
