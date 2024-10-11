import tkinter as tk

root = tk.Tk()
root.title("Table")

# 创建表头
tk.Label(root, text="序号").grid(row=0, column=0)
tk.Label(root, text="检查项目").grid(row=0, column=1)
tk.Label(root, text="结果").grid(row=0, column=2)
tk.Label(root, text="参考值").grid(row=0, column=3)

# 创建结果选项
results = ['阴性（—）', '阳性（—）']

# 创建表格
rows = 8
cols = 4
for i in range(1, rows+1):
    tk.Label(root, text=i).grid(row=i, column=0)
    tk.Entry(root).grid(row=i, column=1)
    tk.StringVar()
    result_option = tk.OptionMenu(root, tk.StringVar(), *results)
    result_option.grid(row=i, column=2)
    tk.Label(root, text='阴性（—）').grid(row=i, column=3)

root.mainloop()