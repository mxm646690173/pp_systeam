import tkinter as tk

print("开始测试 tkinter...")

# 创建主窗口
print("创建主窗口...")
root = tk.Tk()
print(f"主窗口创建成功: {root}")

# 设置窗口标题
root.title("测试窗口")

# 设置窗口大小
root.geometry("300x200")

# 添加标签
label = tk.Label(root, text="Hello, Tkinter!")
label.pack(pady=50)

# 添加按钮
button = tk.Button(root, text="关闭", command=root.destroy)
button.pack()

print("窗口显示完成...")

# 运行主循环
root.mainloop()

print("测试完成...")