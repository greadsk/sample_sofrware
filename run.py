import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import subprocess
import os

def run_script(script_name, params=None, file_path=None):
    """
    运行指定的Python脚本，并传递给定的参数和文件路径。
    
    :param script_name: 要运行的脚本名称（例如 'main_sample.py'）。
    :param params: 要传递给脚本的参数列表。
    :param file_path: Excel文件的路径。
    """
    try:
        # 获取当前脚本所在目录
        current_dir = os.path.dirname(os.path.abspath(__file__))
        # 构造命令以运行脚本并传递参数和文件路径
        command = ['python', script_name]
        if params:
            command += params
        if file_path:
            command.append(file_path)
        
        # 使用subprocess执行脚本
        result = subprocess.run(
            command,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            cwd=current_dir
        )
        
        # 显示成功消息以及脚本输出信息
        messagebox.showinfo("成功", f"{script_name} 执行成功:\n{result.stdout}")
    except subprocess.CalledProcessError as e:
        # 显示错误消息如果脚本执行失败
        messagebox.showerror("错误", f"执行 {script_name} 失败:\n{e.stderr}")

def select_file():
    """
    打开文件对话框选择Excel文件，并更新条目字段。
    """
    file_path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx")])
    if file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(tk.END, file_path)

def run_main_sample():
    """
    收集输入参数和文件路径，然后运行 main_sample.py。
    """
    # 收集所有参数值从下拉表或条目字段
    params = [
        entries[0].get(),
        combos[0].get(),
        combos[1].get(),
        combos[2].get(),
        combos[3].get(),
        combos[4].get(),
        entries[1].get(),
        entries[2].get(),
        combos[5].get(),
        combos[6].get(),
        combos[7].get(),
        combos[8].get(),
        combos[9].get(),
        combos[10].get()
    ]
    file_path = file_entry.get()
    
    # 检查是否所有字段都已填写
    if not all(params + [file_path]):
        messagebox.showwarning("输入错误", "请填写所有字段。")
        return
    
    # 使用收集到的参数和文件路径运行 main_sample.py
    run_script('main_sample.py', params, file_path)

def run_calculate_coefficient():
    """
    收集输入参数和文件路径，然后运行 calculate_coefficient.py。
    """
    # 收集所有参数值从下拉表或条目字段
    params = [
        entries[0].get(),
        combos[0].get(),
        combos[1].get(),
        combos[2].get(),
        combos[3].get(),
        combos[4].get(),
        entries[1].get() or "0",
        entries[2].get() or "0",
        combos[5].get(),
        combos[6].get(),
        combos[7].get(),
        combos[8].get(),
        combos[9].get(),
        combos[10].get()
    ]
    file_path = file_entry.get()
    
    # 检查是否所有字段都已填写
    required_params = params[:6] + params[8:] + [file_path]
    if not all(required_params):
        messagebox.showwarning("输入错误", "请填写所有必填字段。")
        return
    
    # 使用收集到的参数和文件路径运行 calculate_coefficient.py
    run_script('calculate_coefficient.py', params, file_path)

def run_init():
    """
    运行 init.py 程序。
    """
    run_script('init.py')

# 创建主应用程序窗口
root = tk.Tk()
root.title("固定资产投资项目抽样方案设计软件")

# 创建一个框架来容纳所有小部件
frame = tk.Frame(root, padx=10, pady=10)
frame.pack(fill=tk.BOTH, expand=True)

# 定义每个参数的选项
options = [
    [],  # 文件名通常是通过文件选择对话框获取的，这里仅作占位符
    ["项目", "单位"],
    ["行业层", "地级市层"],
    ["行业层", "地级市层"],
    ["等比例分配", "内曼分配","单层抽样"],
    ["计划总投资", "本年完成投资(或自年初累计完成投资)"],
    [],
    [],
    ["None", "计划总投资", "本年完成投资(或自年初累计完成投资)"],
    ["None", "计划总投资", "本年完成投资(或自年初累计完成投资)"],
    [100, 500, 1000, 10000],
    ["中位数", "误差最小值"],
    ["计划总投资", "本年完成投资(或自年初累计完成投资)"],
    ["待开发"]
]

label_name = [
    "数据年月（例如：2403）", "抽样对象", "分层第一层", "分层第二层", "样本分配方式",
    "内曼分配下的方差选择", "抽样比例第一层", "抽样比例第二层", "抽样权重1",
    "抽样权重2", "模拟抽样次数", "抽样结果选择方式", "选择最小误差下的最小化指标", "其他"
]

# 创建标签和输入字段用于每个参数
entries = []
combos = []

for i in range(14):
    row = i // 2
    col = i % 2 * 2
    label_param = tk.Label(frame, text=label_name[i])
    label_param.grid(row=row, column=col, sticky=tk.W, pady=5)
    
    if i == 0:
        entry_param = tk.Entry(frame, width=20)
        entry_param.grid(row=row, column=col+1, pady=5)
        entries.append(entry_param)
    elif i in [6, 7]:
        entry_param = tk.Entry(frame, width=20)
        entry_param.grid(row=row, column=col+1, pady=5)
        entries.append(entry_param)
    else:
        combo_param = ttk.Combobox(frame, values=options[i], state='readonly')
        combo_param.set(options[i][0])  # 设置默认值
        combo_param.grid(row=row, column=col+1, pady=5)
        combos.append(combo_param)

# 创建标签和条目字段用于选择Excel文件
label_file = tk.Label(frame, text="选择Excel文件:")
label_file.grid(row=7, column=0, sticky=tk.W, pady=5)
file_entry = tk.Entry(frame, width=40)
file_entry.grid(row=7, column=1, pady=5)

# 创建按钮以浏览并选择Excel文件
button_select_file = tk.Button(frame, text="浏览...", command=select_file)
button_select_file.grid(row=7, column=2, padx=5, pady=5)

# 创建按钮以运行 init.py
button_run_init = tk.Button(frame, text="第一次运行，安装python相关包", command=run_init)
button_run_init.grid(row=8, column=0, columnspan=3, pady=(20, 10))

# 创建按钮以运行 main_sample.py 和 calculate_coefficient.py
button_run_main_sample = tk.Button(frame, text="运行抽样程序", command=run_main_sample)
button_run_main_sample.grid(row=9, column=0, columnspan=3, pady=(0, 10))

button_run_calculate_coefficient = tk.Button(frame, text="运行入库相关系数计算程序", command=run_calculate_coefficient)
button_run_calculate_coefficient.grid(row=10, column=0, columnspan=3, pady=(0, 10))

# 启动Tkinter事件循环
root.mainloop()



