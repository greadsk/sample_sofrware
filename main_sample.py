
'''
    "#!/usr/bin/env/sample python 3.10.18",
    "# @Time: 2025/07/18 ",
    "# -*- coding: utf-8 -*-",
    "@File: main_simple.py",
    "@Description: 后台主要文件",
    "@software: VsCode",
    "@contact: cqm",
'''

import sys
import pandas as pd
import numpy as np
import os
from tool import * 

# # 输入参数，请在此处修改参数，即可运行
# #文件读入相关参数
# file_name = "2506" #原始数据的文件名，不含后缀，请保证文件后缀为.xlsx

# #抽样相关参数
# #抽样方式，可选项为"项目"、"单位"，按项目进行抽样或者按单位进行抽样
# aim="单位" #"单位"或者"项目"
# #分层方式，第一层是“行业层”，第二层是“地级市层”
# fengcheng1="行业层"
# fengcheng2="地级市层"
# #样本分配方式
# skill="等比例分配" #"等比例分配"或者"内曼分配"或"单层抽样"
# target_col="计划总投资"#按照计划总投资的方差进行内曼分配（等比例分配时不用管）
# #抽样比例：rate1是第一层，即按照行业分层的进行抽样的抽样比例，rate2是第二层，即按照地级市分层的进行抽样的抽样比例
# rate1=0.2
# rate2=0.2
# #抽样权重方式，可选“计划总投资”或者“本年完成投资(或自年初累计完成投资)”
# weight1=None
# weight2=None
# #抽样次数，模拟抽样的总次数
# sample_times=1000 
# #method为抽样结果的选取方式，可以使用按照中位数或者最小误差进行抽样选择
# method="误差最小值" #"中位数"或者"误差最小值"
# choose="计划总投资" # 当method选择“最小误差”时，choose必须有填充，可选“计划总投资”或者“本年完成投资(或自年初累计完成投资)”

# # 输入参数结束

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


#处理输入
if weight1=="None": weight1=None
if weight2=="None": weight2=None
if other=="None": other=None    


#主程序开始运行
os.chdir(os.path.dirname(__file__))

# data = pd.read_excel(os.path.join('data', '{}.xlsx'.format(file_name)))
data= pd.read_excel(file_path)
file_name=file_name





if aim=="单位":
    df=data_preprocessing_danwei(data)
elif aim=="项目":
    df=data_preprocessing_xiangmu(data)
else:
    print("请输入正确的方法")

print("正在抽样中...")
if method=="中位数":
    seed_index=chouyang_median(df,sample_times, skill , target_col, fengcheng1,fengcheng2,rate1,rate2,weight1,weight2)
    print("正在输出结果中...正在计算误差和挑选最优抽样")
    sample_df=output_sample_data(df,file_name,seed_index , skill, target_col,fengcheng1, fengcheng2, rate1, rate2, weight1, weight2)
    each_layer_output(df,sample_df)

elif method=="误差最小值":
    seed_index=chouyang_minerror(df,sample_times, choose,skill,target_col,fengcheng1,fengcheng2,rate1,rate2,weight1,weight2)
    print("正在输出结果中...正在计算误差和挑选最优抽样")
    sample_df=output_sample_data(df,file_name,seed_index,skill,target_col,fengcheng1, fengcheng2, rate1, rate2, weight1, weight2)
    each_layer_output(df,sample_df)


print("抽样完成")

