import sys
import pandas as pd
import numpy as np
import os
from tool import * 

# 假设参数按顺序传递
file_name = sys.argv[1]
aim = sys.argv[2]
fengcheng1 = sys.argv[3]
fengcheng2 = sys.argv[4]
skill = sys.argv[5]
target_col = sys.argv[6]
rate1 = float(sys.argv[7])
rate2 = float(sys.argv[8])
weight1 = sys.argv[9]
weight2 = sys.argv[10]
sample_times = int(sys.argv[11])
method = sys.argv[12]
choose = sys.argv[13]
other = sys.argv[14]
file_path = sys.argv[15]

# 将参数设置为全局变量
globals()['file_name'] = file_name
globals()['aim'] = aim
globals()['fengcheng1'] = fengcheng1
globals()['fengcheng2'] = fengcheng2
globals()['skill'] = skill
globals()['target_col'] = target_col
globals()['rate1'] = rate1
globals()['rate2'] = rate2
globals()['weight1'] = weight1
globals()['weight2'] = weight2
globals()['sample_times'] = sample_times
globals()['method'] = method
globals()['choose'] = choose
globals()['other'] = other
globals()['file_path'] = file_path


# print("新入库项目计划投资系数计算程序开始运行,请稍等...")
# data = pd.read_excel(os.path.join('data', '{}.xlsx'.format(file_name)))
# df=data_preprocessing_danwei(data)
# new_inventory_ratio(df,file_name)


# # 输入参数，请在此处修改参数，即可运行
# #文件读入相关参数
# file_name = "2506" #原始数据的文件名，不含后缀，请保证文件后缀为.xlsx

# #输出参数结束

#主程序开始运行
print("新入库项目计划投资系数计算程序开始运行,请稍等...")
os.chdir(os.path.dirname(__file__))
# data = pd.read_excel(os.path.join('data', '{}.xlsx'.format(file_name)))
data= pd.read_excel(file_path)
file_name=file_name
df=data_preprocessing_xiangmu(data)
ls=new_inventory_ratio(df,file_name)
ls.to_excel('新入库系数.xlsx', index=False)

print("入库时间: ",ls['入库时间'].iloc[0])
print("新入库项目计划总投资系数:",ls["新入库项目计划总投资系数"].iloc[0])
print("新入库项目本年完成投资系数:",ls["新入库项目本年完成投资系数"].iloc[0])
print("新入库系数计算完毕,保存在'新入库系数.xlsx'中")




# #临时文件
# os.chdir(os.path.dirname(__file__))
# year=["2303","2304","2305","2306","2403","2404","2405","2406","2407","2408","2409","2410","2411","2412","2503","2504","2505","2506"]
# go=pd.DataFrame()
# for file_name in year:
#     print("新入库项目计划投资系数计算程序开始运行,请稍等...")
#     data = pd.read_excel(os.path.join('data', '{}.xlsx'.format(file_name)))
#     df=data_preprocessing_xiangmu(data)
#     now=new_inventory_ratio(df,file_name)
#     go=pd.concat([go,now], axis=0, ignore_index=True)
# go.to_excel(os.path.join('{}.xlsx'.format("new_inventory_ratio")))


# #临时文件,按单位来抽
# os.chdir(os.path.dirname(__file__))
# year=["2303","2304","2305","2306","2403","2404","2405","2406","2407","2408","2409","2410","2411","2412","2503","2504","2505","2506"]
# go=pd.DataFrame()
# for file_name in year:
#     print("新入库项目计划投资系数计算程序开始运行,请稍等...")
#     data = pd.read_excel(os.path.join('data', '{}.xlsx'.format(file_name)))
#     df=data_preprocessing_danwei(data)
#     now=new_inventory_ratio(df,file_name)
#     go=pd.concat([go,now], axis=0, ignore_index=True)
# go.to_excel(os.path.join('{}.xlsx'.format("new_inventory_ratio")))