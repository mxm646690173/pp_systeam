import tkinter as tk
from tkinter import ttk
import warnings
import requests
import json
import time
import os
import logging
import sys

VERSION = "1.0.0"
APP_NAME = "毛坯采购分析工具"

# 配置日志（使用安全的路径处理）
try:
    # 使用当前工作目录创建日志文件路径
    log_file = os.path.join(os.getcwd(), 'app.log')
    
    # 配置日志，只使用StreamHandler，避免文件写入问题
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler()
        ]
    )
    
    # 尝试添加文件处理器
    try:
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.INFO)
        file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logging.getLogger().addHandler(file_handler)
    except Exception as e:
        # 如果文件处理器创建失败，只使用流处理器
        logging.warning(f"无法创建日志文件处理器: {e}")
        logging.warning("将只使用控制台输出")
except Exception as e:
    # 如果日志配置完全失败，使用默认配置
    logging.basicConfig(level=logging.INFO)
    logging.error(f"日志配置失败: {e}")

# 禁用所有警告
warnings.filterwarnings('ignore')

# 禁用 HTTPS 警告（当使用 HTTP 协议时）
try:
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
except ImportError:
    pass

# 禁用 requests 库的 HTTPS 警告
try:
    from requests.packages.urllib3.exceptions import InsecureRequestWarning
    requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
except ImportError:
    pass

# 导入拆分后的模块
from ui_modules import build_main_interface
from data_handling import init_data_folders, load_history, save_history, save_token

# 定义TOKEN文件路径
TOKEN_FILE = os.path.join(os.getcwd(), 'token.txt')

# 默认路径配置
PATH_CONFIG = {
    "base_data": "./基础数据",
    "daily_upload": "./每日上传",
    "report": "./分析报告",
    "backup": "./Backup",
    "history": "./history.json"
}

# 全局变量（存储所有数据）
global_data = {
    "raw_material_df": None,    # 毛坯基础信息
    "weekly_order_df": None,    # 每周订单
    "finished_stock_df": None,  # 当日成品库存
    "raw_stock_df": None,       # 当日毛坯库存
    "analysis_result_df": None, # 分析结果
    "last_import_paths": {}     # 存储上次导入的文件路径
}

# 直接在主程序中打开嵌入式登录页面并获取token
def open_embedded_login():
    """
    直接在主程序中打开嵌入式登录页面并获取token（使用playwright）
    """
    print("进入 open_embedded_login 函数...")
    
    # 初始化变量
    token = None
    startup_window = None
    browser_closed = False
    
    try:
        from playwright.sync_api import sync_playwright
        import time
        
        # 登录网页URL（确保使用 HTTP 协议）
        login_url = "http://61.153.238.198:9331/#/login"
        
        # 优化URL处理，确保使用正确的协议和地址
        if not login_url.startswith("http://") and not login_url.startswith("https://"):
            login_url = "http://" + login_url
        
        # 创建启动提示弹窗
        print("创建启动提示弹窗...")
        try:
            startup_window = tk.Tk()
            startup_window.title("启动中")
            startup_window.geometry("300x100")
            startup_window.resizable(False, False)
            
            # 居中显示
            screen_width = startup_window.winfo_screenwidth()
            screen_height = startup_window.winfo_screenheight()
            x = (screen_width - 300) // 2
            y = (screen_height - 100) // 2
            startup_window.geometry(f"300x100+{x}+{y}")
            
            # 添加提示标签
            label = tk.Label(startup_window, text="程序启动中...\n请稍候，正在打开浏览器", font=("微软雅黑", 12))
            label.pack(expand=True)
            
            # 更新窗口显示
            startup_window.update()
        except Exception as e:
            print(f"创建启动提示弹窗时出错: {e}")
        
        # 创建并显示浏览器窗口
        print("创建嵌入式浏览器窗口（使用playwright）...")
        
        # 尝试查找playwright的浏览器驱动
        import os
        import sys
        playwright_browser_path = None
        
        # 检查系统中安装的 Chrome 浏览器
        # 首先检查 Playwright 安装的浏览器
        ms_playwright_path = os.path.expanduser('~/AppData/Local/ms-playwright')
        if os.path.exists(ms_playwright_path):
            for root, dirs, files in os.walk(ms_playwright_path):
                if 'chrome.exe' in files:
                    playwright_browser_path = os.path.join(root, 'chrome.exe')
                    print(f"找到 Playwright 安装的 Chrome 浏览器: {playwright_browser_path}")
                    break
        
        # 如果没有找到 Playwright 安装的浏览器，尝试查找系统中安装的 Chrome
        if not playwright_browser_path:
            # 检查常见的 Chrome 安装路径
            chrome_paths = [
                os.path.join(os.environ.get('ProgramFiles', 'C:/Program Files'), 'Google', 'Chrome', 'Application', 'chrome.exe'),
                os.path.join(os.environ.get('ProgramFiles(x86)', 'C:/Program Files (x86)'), 'Google', 'Chrome', 'Application', 'chrome.exe'),
            ]
            
            for path in chrome_paths:
                if os.path.exists(path):
                    playwright_browser_path = path
                    print(f"找到系统安装的 Chrome 浏览器: {playwright_browser_path}")
                    break
        
        # 如果仍然没有找到，让 Playwright 使用默认的浏览器
        if not playwright_browser_path:
            print("未找到 Chrome 浏览器，使用 Playwright 默认浏览器")
        
        with sync_playwright() as p:
            # 启动Chromium浏览器（进一步优化启动速度）
            browser = p.chromium.launch(
                headless=False,  # 显示浏览器窗口
                executable_path=playwright_browser_path,  # 指定浏览器路径
                args=[
                    '--window-size=800,600',
                    '--disable-infobars',
                    '--disable-extensions',
                    '--disable-gpu',
                    '--no-sandbox',
                    '--disable-dev-shm-usage',
                    '--disable-software-rasterizer',
                    '--disable-background-networking',
                    '--disable-sync',
                    '--disable-translate',
                    '--metrics-recording-only',
                    '--no-first-run',
                    '--safebrowsing-disable-auto-update',
                    '--disable-crash-reporter',
                    '--disable-logging',
                    '--disable-breakpad',
                    '--mute-audio'
                ]
            )
            
            # 创建浏览器上下文
            context = browser.new_context(
                viewport={'width': 800, 'height': 600},
                user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            )
            
            # 创建新页面
            page = context.new_page()
            
            # 关闭启动提示弹窗
            print("关闭启动提示弹窗...")
            try:
                if startup_window:
                    startup_window.destroy()
                    startup_window = None
            except Exception as e:
                print(f"关闭启动提示弹窗时出错: {e}")
            
            print(f"打开登录页面: {login_url}")
            page.goto(login_url, wait_until='domcontentloaded', timeout=30000)
            
            print("登录网页已加载")
            print("请在嵌入式浏览器中输入账号密码和验证码进行登录")
            
            # 账号和密码
            username = "03001"
            password = "88981788"
            
            # 尝试填充账号
            try:
                print("尝试填充账号...")
                # 使用playwright的fill方法
                page.locator('input').first.fill(username)
                print(f"已填充账号: {username}")
            except Exception as e:
                print(f"填充账号时出错: {e}")
            
            # 尝试填充密码
            try:
                print("尝试填充密码...")
                # 使用playwright的fill方法
                page.locator('input').nth(1).fill(password)
                print(f"已填充密码: {password}")
            except Exception as e:
                print(f"填充密码时出错: {e}")
            
            # 尝试设置验证码输入框焦点
            try:
                print("尝试设置验证码输入框焦点...")
                page.locator('input').nth(2).focus()
                print("已设置验证码输入框焦点")
            except Exception as e:
                print(f"设置验证码输入框焦点时出错: {e}")
            
            # 尝试将网页滚动到最右侧
            try:
                print("尝试将网页滚动到最右侧...")
                page.evaluate('window.scrollTo(99999, 0)')
                print("已滚动网页到最右侧")
            except Exception as e:
                print(f"滚动网页时出错: {e}")
            
            # 填充操作完成
            print("填充操作完成")
            
            # 定期检查是否获取到token（优化检查间隔）
            print("开始定期检查token...")
            max_wait_time = 300  # 最多等待5分钟
            start_time = time.time()
            browser_closed = False  # 初始化浏览器关闭状态
            
            print(f"初始化完成，browser_closed={browser_closed}")
            
            while time.time() - start_time < max_wait_time and not browser_closed:
                print(f"循环开始，browser_closed={browser_closed}")
                time.sleep(1)  # 缩短检查间隔到1秒
                
                # 先检查浏览器是否仍然打开
                try:
                    print("检查浏览器连接状态...")
                    if not browser.is_connected():
                        print("浏览器已关闭")
                        browser_closed = True
                        print(f"设置browser_closed={browser_closed}")
                        # 弹出对话框提示用户
                        from tkinter import messagebox
                        messagebox.showinfo("提示", "主动关闭浏览器，操作被取消")
                        print("弹出提示对话框，准备退出循环")
                        break
                except Exception as e:
                    print(f"无法检查浏览器状态，假设浏览器已关闭: {e}")
                    browser_closed = True
                    print(f"设置browser_closed={browser_closed}")
                    # 弹出对话框提示用户
                    try:
                        from tkinter import messagebox
                        messagebox.showinfo("提示", "主动关闭浏览器，操作被取消")
                        print("弹出提示对话框，准备退出循环")
                    except Exception as msg_error:
                        print(f"弹出对话框失败: {msg_error}")
                    break
                
                # 尝试获取token
                try:
                    print("尝试获取token...")
                    # 尝试从浏览器的localStorage中获取token
                    token_value = page.evaluate('localStorage.getItem("token")')
                    if token_value:
                        token = token_value
                        print(f"自动获取到token: {token}")
                        # 保存token
                        save_token(token)
                        print("token已保存")
                        break
                    
                    # 尝试从浏览器的sessionStorage中获取token
                    token_value = page.evaluate('sessionStorage.getItem("token")')
                    if token_value:
                        token = token_value
                        print(f"自动获取到token: {token}")
                        # 保存token
                        save_token(token)
                        print("token已保存")
                        break
                    
                    # 尝试从浏览器的cookies中获取token
                    cookies = context.cookies()
                    for cookie in cookies:
                        if cookie['name'] == 'token':
                            token = cookie['value']
                            print(f"自动获取到token: {token}")
                            # 保存token
                            save_token(token)
                            print("token已保存")
                            break
                    
                    if token:
                        break
                        
                except Exception as e:
                    # 如果获取失败，检查是否是浏览器关闭导致的
                    error_msg = str(e)
                    print(f"捕获到异常: {error_msg}")
                    if "Target page, context or browser has been closed" in error_msg:
                        print("浏览器已关闭，停止获取token")
                        browser_closed = True
                        print(f"设置browser_closed={browser_closed}")
                        # 不再弹出对话框提示用户，因为在浏览器连接状态检查时已经弹出过了
                        break
                    else:
                        print(f"获取token时出错: {e}")
                        # 检查浏览器是否仍然打开
                        try:
                            print("检查浏览器连接状态...")
                            if not browser.is_connected():
                                print("浏览器已关闭")
                                browser_closed = True
                                print(f"设置browser_closed={browser_closed}")
                                # 不再弹出对话框提示用户，因为在浏览器连接状态检查时已经弹出过了
                                break
                        except Exception as conn_error:
                            print(f"无法检查浏览器状态，假设浏览器已关闭: {conn_error}")
                            browser_closed = True
                            print(f"设置browser_closed={browser_closed}")
                            # 不再弹出对话框提示用户，因为在浏览器连接状态检查时已经弹出过了
                            break
                    # 如果不是浏览器关闭导致的错误，继续尝试
                    continue
            
            # 关闭浏览器
            try:
                if browser.is_connected():
                    browser.close()
            except Exception as e:
                print(f"关闭浏览器时出错: {e}")
            
    except Exception as e:
        print(f"打开嵌入式登录页面时出错: {e}")
        import traceback
        traceback.print_exc()
        
        # 尝试关闭启动提示弹窗
        try:
            if startup_window:
                startup_window.destroy()
        except:
            pass
        
        # 浏览器关闭后，检查是否获取到token
        if not token:
            # 如果浏览器已关闭，直接返回False，不弹出让用户输入token的弹窗
            if browser_closed:
                return False
            # 如果自动获取失败，提示用户手动输入
            try:
                # 使用tkinter的simpledialog代替input，避免在GUI程序中使用命令行输入
                from tkinter import simpledialog, Tk
                
                # 创建临时根窗口
                root = Tk()
                root.withdraw()  # 隐藏主窗口
                
                # 弹出输入对话框
                token_input = simpledialog.askstring(
                    "Token输入",
                    "浏览器已关闭，但未获取到token，请手动输入token:",
                    parent=root
                )
                
                # 销毁临时窗口
                root.destroy()
                
                if token_input:
                    token = token_input
                    print(f"获取到token: {token}")
                    # 保存token
                    save_token(token)
                    print("token已保存")
                else:
                    print("未输入token，登录失败")
                    return False
            except Exception as e:
                print(f"获取token失败: {e}")
                return False
    finally:
        # 确保启动提示弹窗被关闭
        try:
            if startup_window:
                startup_window.destroy()
        except:
            pass
    
    # 浏览器窗口关闭后，检查是否获取到token
    if not token:
        # 如果浏览器已关闭，直接返回False，不弹出让用户输入token的弹窗
        if browser_closed:
            print("浏览器已关闭，登录失败")
            return False
        # 如果自动获取失败，提示用户手动输入
        try:
            # 使用tkinter的simpledialog代替input，避免在GUI程序中使用命令行输入
            from tkinter import simpledialog, Tk
            
            # 创建临时根窗口
            root = Tk()
            root.withdraw()  # 隐藏主窗口
            
            # 弹出输入对话框
            token_input = simpledialog.askstring(
                "Token输入",
                "浏览器已关闭，但未获取到token，请手动输入token:",
                parent=root
            )
            
            # 销毁临时窗口
            root.destroy()
            
            if token_input:
                token = token_input
                print(f"获取到token: {token}")
                # 保存token
                save_token(token)
                print("token已保存")
            else:
                print("未输入token，登录失败")
                return False
        except Exception as e:
            print(f"获取token失败: {e}")
            return False

    return True

# 主函数
def main():
    """
    主函数
    """
    print("进入 main 函数...")
    
    # 初始化默认文件夹
    init_data_folders(PATH_CONFIG)
    
    # 加载历史导入记录
    global_data["last_import_paths"] = load_history(PATH_CONFIG)
    
    print("直接启动主程序...")
    
    # 创建主窗口
    root = tk.Tk()
    root.title(f"{APP_NAME} v{VERSION}")
    root.geometry("1200x800")
    root.resizable(True, True)  # 允许窗口最大化
    
    # 优化文字清晰度（适用于高DPI屏幕）
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
        root.tk.call('tk', 'scaling', 1.2)  # 适当调整缩放比例
    except Exception as e:
        # print(f"DPI优化失败: {e}")
        pass
    
    # 美化样式
    style = ttk.Style(root)
    style.configure("Accent.TButton", font=("微软雅黑", 10, "bold"))
    
    # 构建主界面（传入全局数据和路径配置）
    build_main_interface(root, global_data, PATH_CONFIG)
    
    # 添加关闭事件处理函数，保存历史记录
    def on_closing():
        save_history(global_data, PATH_CONFIG)
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    # 运行主循环
    root.mainloop()

if __name__ == "__main__":
    print("程序开始运行...")
    main()