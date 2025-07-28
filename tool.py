import pandas as pd
import numpy as np
import os


# 处理数据（按照单位进行抽样）
def data_preprocessing_danwei(df): 
    # 定义处理行业代码的函数
    def process_industry_code(code):
        code_str = str(code)  # 确保code是字符串类型
        if len(code_str) == 3:
            return code_str[0]
        elif len(code_str) >= 4:
            return code_str[:2]
        else:
            return code_str  # 如果长度不符合条件，则返回原值或进行其他处理
    # 数据清洗和合并（按单位抽）
    df['建设地址'] = df['建设地址'].astype(str)  # 将“建设地址”列转换为字符串类型
    df['法人单位码'] = df['法人单位码'].astype(str) 
    df = df.groupby('法人单位码').agg({
        '计划总投资': 'sum',
        '本年完成投资(或自年初累计完成投资)': 'sum',
        '行业代码2017': lambda x: x.apply(process_industry_code).iloc[0],
        '建设地址': lambda x: x.str[:4].iloc[0],
        '投资入库时间': 'max' #"first" or "max"
    }).reset_index()
    # 确保“行业代码2017”列是整数类型
    df['行业代码2017'] = df['行业代码2017'].astype(str).str.extract('(\d+)').astype(float).astype(pd.Int64Dtype())
    df['行业代码2017'] = pd.to_numeric(df['行业代码2017'], errors='coerce').fillna(0).astype(int)
    # 定义行业分层映射
    industry_mapping = {
        '农林牧渔': (1, 5), '采矿业': (6, 12), '制造业': (13, 43),
        '电力热力燃气及水的生产和供应业': (44, 46), '建筑业': (47, 50), '批发零售业': (51, 52),
        '交通运输、仓储和邮政业': (53, 60), '住宿餐饮业': (61, 62), '信息传输、软件和信息技术业': (63, 65),
        '金融业': (66, 69), '房地产业': (70, 70), '租赁和商务服务业': (71, 72),
        '科学技术研究和技术服务业': (73, 75), '水利、环境和公共设施治理业': (76, 79),
        '居民服务、修理和其他服务业': (80, 82), '教育': (83, 83), '卫生社会工作': (84, 85),
        '文化体育和娱乐业': (86, 90), '公共管理社会保障和社会组织': (91, 97)
    }
    # 按照行业代码2017进行第一层分层
    def assign_industry_layer(code):
        for layer, (start, end) in industry_mapping.items():
            if start <= code <= end:
                return layer
        return '其他'
    df['行业层'] = df['行业代码2017'].apply(assign_industry_layer)
    # 假设“建设地址”列中包含了地级市信息，并且已经为数字形式
    # 如果不是，请先转换为相应的编码或数字表示
    df['地级市层'] = df['建设地址']
    selected_columns = df.filter(['行业层', '地级市层', '法人单位码', '计划总投资', '本年完成投资(或自年初累计完成投资)',
                             '投资入库时间'])
    return selected_columns 


#按项目进行抽样
def data_preprocessing_xiangmu(df):
# 定义处理行业代码的函数
    def process_industry_code(code):
        code_str = str(code)  # 确保code是字符串类型
        if len(code_str) == 3:
            return code_str[0]
        elif len(code_str) >= 4:
            return code_str[:2]
        else:
            return code_str  # 如果长度不符合条件，则返回原值或进行其他处理

    # 数据清洗和合并（按单位抽）
    df['建设地址'] = df['建设地址'].astype(str)  # 将“建设地址”列转换为字符串类型
    df = df.groupby('项目(法人)码').agg({
        '计划总投资': 'sum',
        '本年完成投资(或自年初累计完成投资)': 'sum',
        '行业代码2017': lambda x: x.apply(process_industry_code).iloc[0],
        '建设地址': lambda x: x.str[:4].iloc[0],
        '投资入库时间': 'first'
    }).reset_index()
    # 确保“行业代码2017”列是整数类型
    df['行业代码2017'] = df['行业代码2017'].astype(str).str.extract('(\d+)').astype(float).astype(pd.Int64Dtype())
    df['行业代码2017'] = pd.to_numeric(df['行业代码2017'], errors='coerce').fillna(0).astype(int)
    # 定义行业分层映射
    industry_mapping = {
        '农林牧渔': (1, 5), '采矿业': (6, 12), '制造业': (13, 43),
        '电力热力燃气及水的生产和供应业': (44, 46), '建筑业': (47, 50), '批发零售业': (51, 52),
        '交通运输、仓储和邮政业': (53, 60), '住宿餐饮业': (61, 62), '信息传输、软件和信息技术业': (63, 65),
        '金融业': (66, 69), '房地产业': (70, 70), '租赁和商务服务业': (71, 72),
        '科学技术研究和技术服务业': (73, 75), '水利、环境和公共设施治理业': (76, 79),
        '居民服务、修理和其他服务业': (80, 82), '教育': (83, 83), '卫生社会工作': (84, 85),
        '文化体育和娱乐业': (86, 90), '公共管理社会保障和社会组织': (91, 97)
    }
    # 按照行业代码2017进行第一层分层
    def assign_industry_layer(code):
        for layer, (start, end) in industry_mapping.items():
            if start <= code <= end:
                return layer
        return '其他'
    df['行业层'] = df['行业代码2017'].apply(assign_industry_layer)
    # 假设“建设地址”列中包含了地级市信息，并且已经为数字形式
    # 如果不是，请先转换为相应的编码或数字表示
    df['地级市层'] = df['建设地址']
    selected_columns = df.filter(['行业层', '地级市层', '项目(法人)码', '计划总投资', '本年完成投资(或自年初累计完成投资)',
                             '投资入库时间'])
    return selected_columns 






# 计算新入库项目计划总投资系数
def new_inventory_ratio(df,file_name):
    time="20"+str(file_name)
    filtered_df = df[df['投资入库时间'] == int(time)]
    total_new_investment = filtered_df['计划总投资'].sum()
    total_new_investment_plan=filtered_df['本年完成投资(或自年初累计完成投资)'].sum()
    total=df['计划总投资'].sum()
    total_plan=df['本年完成投资(或自年初累计完成投资)'].sum()
    ratio=1/(1-total_new_investment/total)
    ratio_plan=1/(1-total_new_investment_plan/total_plan)


    #计算总和
    total_plan = df['计划总投资'].sum()
    total_complete = df['本年完成投资(或自年初累计完成投资)'].sum()

    #计算均值，缺失值用0填补来计算均值
    mean_plan = df['计划总投资'].fillna(0).mean()
    mean_complete = df['本年完成投资(或自年初累计完成投资)'].fillna(0).mean()

    #计算方差
    variance_plan = df['计划总投资'].var()
    variance_complete = df['本年完成投资(或自年初累计完成投资)'].fillna(0).var()


    ls=pd.DataFrame({"入库时间":int(time),
                "新入库项目计划总投资系数":ratio,
                "新入库项目本年完成投资系数":ratio_plan,
                "计划总投资": total_plan,
                "本年完成投资(或自年初累计完成投资)总和":total_complete ,
                '计划总投资均值':mean_plan,
                '本年完成投资均值':mean_complete,
                '计划总投资方差':variance_plan,
                '本年完成投资方差':variance_complete
                },index=[0])
    # ls.to_excel('新入库系数.xlsx', index=False)
    # print("入库时间: ",int(time))
    # print("新入库项目计划总投资系数:",ratio)
    # print("新入库项目本年完成投资系数:",ratio_plan)
    # print("新入库系数计算完毕")
    return ls



#抽样函数

#分层抽样函数
#等比例抽样
def fencheng_sampling(df, stratify_col, frac , weights=None ,random_state=None):
    """
    分层抽样函数
    :param df: 待抽样的 DataFrame
    :param stratify_col: 分层列名:"行业层"、"地级市层"
    :param frac: 抽样比例
    :param weights: 权重列名
    :return: 抽样结果
    """
    # 获取所有唯一的“层”标签
    strata = df[stratify_col].unique()
    # 初始化一个空的 DataFrame 用于保存抽样结果
    sampled_df = pd.DataFrame()
    # 对每一个“层”进行抽样
    for stratum in strata:
        # 提取当前层的数据
        stratum_df = df[df[stratify_col] == stratum]
        # 计算当前层应抽取的样本数量
        sample_size = max(int(round(len(stratum_df) * frac)), 1)
        # 抽样：如果样本数不足 1，则强制抽取 1 条
        stratum_sample = stratum_df.sample(n=sample_size, weights=weights ,random_state=random_state)
        # 将当前层的抽样结果加入最终结果中
        sampled_df = pd.concat([sampled_df, stratum_sample], axis=0)
    # 返回抽样后的 DataFrame
    return sampled_df

#内曼分配
def neyman_sampling(df, target_col , stratify_col, frac, weights=None,random_state=None):
    """
    使用 Neyman 分配的分层抽样函数，基于抽样比例 frac
    :param df: 待抽样的 DataFrame
    :param stratify_col: 分层列名，例如 "行业层"、"地级市层"
    :param target_col: 目标变量列名，用于计算方差
    :param frac: 抽样比例（总样本数 = 总体样本数 × frac）
    :param random_state: 随机种子
    :return: 抽样结果 DataFrame
    """
    # 获取所有唯一的“层”标签
    strata = df[stratify_col].unique()
    # 初始化一个空的 DataFrame 用于保存抽样结果
    sampled_df = pd.DataFrame()
    # 总体样本数
    total_n = len(df)
    # 总抽样样本数
    total_sample_size = max(int(round(total_n * frac)), 1)

    # 计算每层的 Nh * Sh
    total_weight = 0
    stratum_info = {}

    for stratum in strata:
        stratum_df = df[df[stratify_col] == stratum]
        if len(stratum_df) == 0:
            continue  # 跳过空层
        N_h = len(stratum_df)
        S_h = stratum_df[target_col].std()  # 标准差
        if pd.isna(S_h) or S_h == 0:
            S_h = 1e-6  # 给一个很小的默认值，防止权重为 0
        weight = N_h * S_h
        stratum_info[stratum] = {'N_h': N_h, 'S_h': S_h, 'weight': weight}
        total_weight += weight
    if total_weight == 0:
        print("所有层的方差均为 0，采用等比例抽样")
        for stratum in strata:
            stratum_df = df[df[stratify_col] == stratum]
            n_h = max(int(round(len(stratum_df) * frac)), 1)
            stratum_sample = stratum_df.sample(n=n_h, random_state=random_state)
            sampled_df = pd.concat([sampled_df, stratum_sample], axis=0)
        return sampled_df

    # 按 Neyman 分配每层样本数
    for stratum, info in stratum_info.items():
        n_h = round(total_sample_size * (info['weight'] / total_weight))
        # 至少取一个样本
        n_h = max(n_h, 1)
        stratum_df = df[df[stratify_col] == stratum]
        # 随机抽样（不超过该层总数）
        stratum_sample = stratum_df.sample(n=min(n_h, len(stratum_df)), weights=weights,random_state=random_state)
        sampled_df = pd.concat([sampled_df, stratum_sample], axis=0)

    return sampled_df



#多层抽样函数

def chouyangfangan(df ,skill=None,target_col=None, fengcheng1=None,fengcheng2=None,rate1=None,rate2=None,weight1=None,weight2=None,seed=None):
    """
    :param df: 输入数据框
    :param fengcheng1: 行业层
    :param fengcheng2: 地级市层
    :param rate1: 行业层采样比例
    :param rate2: 地级市层采样比例
    :param weight1: 行业层权重
    :param weight2: 地级市层权重
    :seed : 抽样随机数种子
    :return: 采样后的数据
    """
    if skill==None:
        print("请输入分配方式参数")
        return 
    elif skill=="等比例分配":
        sample=fencheng_sampling(df,fengcheng1,rate1,weight1,random_state=seed)
        sample_df=fencheng_sampling(sample,fengcheng2,rate2,weight2,random_state=seed)
        return sample_df
    elif skill=="加权分配":
        print("可以调整等比例分配的参数实现根据数量大小进行加权分配")
        return 
    elif skill=="内曼分配":
        sample=neyman_sampling(df, target_col, fengcheng1,rate1,weight1,random_state=seed)
        sample_df=neyman_sampling(sample, target_col ,fengcheng2,rate2,weight2,random_state=seed)
        return sample_df
    elif skill=="单层抽样":
        sample=fencheng_sampling(df,fengcheng1,rate1,weight1,random_state=seed)
        return sample


#按照行业再按照地区进行多层按数量等比例抽样

#按照中位数进行抽取
def chouyang_median(df,sample_times, skill, target_col, fengcheng1,fengcheng2,rate1,rate2,weight1,weight2=None):
    if fengcheng1==None or fengcheng2==None or rate1==None or rate2==None or sample_times==None or skill==None or target_col==None:
        print("缺少参数无法运行！请输入参数")
        return 
    def calculate_total(df,col="计划总投资"):
        return df[col].sum()/(rate1*rate2)
    def calculate_total_actual(df,col="计划总投资"):
        return df[col].sum()
    #计算实际总额
    total_actual=calculate_total_actual(df,"计划总投资")
    complete_actual=calculate_total_actual(df,"本年完成投资(或自年初累计完成投资)")
    #抽样模拟
    ls_total=[];ls_complete=[]
    for i in range(sample_times):
        sample_df=chouyangfangan(df ,skill,target_col, fengcheng1,fengcheng2,rate1,rate2,weight1,weight2,seed=i)
        ls_total.append(calculate_total(sample_df,"计划总投资"))
        ls_complete.append(calculate_total(sample_df,"本年完成投资(或自年初累计完成投资)"))
    ls_total=np.array(ls_total)
    ls_complete=np.array(ls_complete)
    median_value=np.median(ls_total)
    median_index=np.isclose(ls_total, median_value).argmax()
    # 输出 抽样种子
    return median_index

#通过最小化误差来抽取样本，choose参数是选择用"计划总投资"最小还是"本年完成投资(或自年初累计完成投资)"最小来抽取样本
def chouyang_minerror(df,sample_times, choose ,skill,target_col, fengcheng1,fengcheng2,rate1,rate2,weight1,weight2):
    if fengcheng1==None or fengcheng2==None or rate1==None or rate2==None or sample_times==None or choose==None or skill==None or target_col==None:
        print("缺少参数无法运行！请输入参数")
        return 
    def calculate_total(df,col="计划总投资"):
        return df[col].sum()/(rate1*rate2)
    def calculate_total_actual(df,col="计划总投资"):
        return df[col].sum()
    #计算实际总额
    total_actual=calculate_total_actual(df,"计划总投资")
    complete_actual=calculate_total_actual(df,"本年完成投资(或自年初累计完成投资)")
    #抽样模拟
    ls_total=[];ls_complete=[]
    for i in range(sample_times):
        sample_df=chouyangfangan(df,skill,target_col, fengcheng1,fengcheng2,rate1,rate2,weight1,weight2,seed=i)
        ls_total.append(abs(calculate_total(sample_df,"计划总投资")-total_actual))
        ls_complete.append(abs(calculate_total(sample_df,"本年完成投资(或自年初累计完成投资)")-complete_actual))
    #输出抽样种子数
    ls_total=np.array(ls_total)
    ls_complete=np.array(ls_complete)
    if choose=="计划总投资":
        min_value=min(ls_total)
        min_index=np.isclose(ls_total, min_value).argmax()
        return min_index
    elif choose=="本年完成投资(或自年初累计完成投资)":
        min_value=min(ls_complete)
        min_index=np.isclose(ls_total, min_value).argmax()
        return min_index




# 输出抽样的结果
def output_sample_data(df,file_name,seed_index,skill,target_col, fengcheng1,fengcheng2,rate1,rate2,weight1,weight2):
    # global  fengcheng1,fengcheng2,rate1,rate2,weight1,weight2
    #抽取样本
    sample_df=chouyangfangan(df, skill,target_col, fengcheng1,fengcheng2,rate1,rate2,weight1,weight2, seed=seed_index)
    def calculate_total(df,col):
        return df[col].sum()/rate1/rate2
    def calculate_total_actual(df,col):
        return df[col].sum()
    #计算实际总额
    total_actual=calculate_total_actual(df,"计划总投资")
    complete_actual=calculate_total_actual(df,"本年完成投资(或自年初累计完成投资)")
    #计算均值、方差、误差等其他指标
    mean_plan_actual=df["计划总投资"].mean()
    variance_plan_actual=df["计划总投资"].var()
    mean_complete_actual=df["本年完成投资(或自年初累计完成投资)"].fillna(0).mean()
    variance_complete_actual=df["本年完成投资(或自年初累计完成投资)"].fillna(0).var()
    #计算抽样总额
    total_sample_plan=calculate_total(sample_df,"计划总投资")
    complete_sample_plan=calculate_total(sample_df,"本年完成投资(或自年初累计完成投资)")
    #计算均值、方差、误差等其他指标
    mean_plan=sample_df["计划总投资"].mean()
    mean_complete=sample_df["本年完成投资(或自年初累计完成投资)"].fillna(0).mean()
    variance_plan=sample_df["计划总投资"].var()
    variance_complete=sample_df["本年完成投资(或自年初累计完成投资)"].fillna(0).var()
    error_plan=abs(total_actual-total_sample_plan)/total_actual
    error_complete=abs(complete_actual-complete_sample_plan)/complete_actual

    chouyang=pd.DataFrame({"实际计划总投资":total_actual,
                        "实际本年完成投资": complete_actual,
                        "抽样得到计划总投资":total_sample_plan,
                        "抽样得到本年完成投资":complete_sample_plan,
                        "计划总投资误差":error_plan,
                        "完成投资误差":error_complete,
                        "实际总投资均值":mean_plan_actual,
                        "实际本年完成投资均值":mean_complete_actual,
                        "抽样总投资均值":mean_plan,
                        "抽样本年完成投资均值":mean_complete,
                        "实际总投资方差" :variance_plan_actual,
                        "实际本年完成投资方差":variance_complete_actual,
                        "抽样总投资方差":variance_plan,
                        "抽样本年完成投资方差":variance_complete
                        }, index=[0])
    cur=pd.DataFrame({"时间":int("20"+str(file_name))},index=[0])
    outcome=pd.concat([cur,chouyang],axis=1)
    outcome.to_excel('抽样计算结果.xlsx', index=False)
    sample_df.to_excel('入样的样本.xlsx', index=False)#可以通过唯一识别码结合excel的vlookup函数找到该样本
    print("抽样计算结果已保存在当前目录下")
    print("入样的样本已保存在当前目录下")
    return sample_df
    
    
    
    

def each_layer_output(df,sample_df):
    """
    对每一层进行统计
    """
    # 统计总个数
    industry_code_counts_total = df['行业层'].value_counts().reset_index()
    industry_code_counts_total.columns = ['行业层', '总数']
    # 统计入样个数
    industry_code_counts = sample_df['行业层'].value_counts().reset_index()
    industry_code_counts.columns = ['行业层', '入样数']
    # 合并总表和分表
    hangye_merged_df = pd.merge(industry_code_counts, industry_code_counts_total, on='行业层')
    # 计算入样比例
    hangye_merged_df['入样比例'] = hangye_merged_df['入样数'] / hangye_merged_df['总数']


    # 按地级市分类进行输出
    city_code_counts_total = df['地级市层'].value_counts().reset_index()
    city_code_counts_total.columns = ['地级市层', '总数']
    city_code_counts = sample_df['地级市层'].value_counts().reset_index()
    city_code_counts.columns = ['地级市层', '入样数']
    #合并表格
    city_merged_df = pd.merge(city_code_counts, city_code_counts_total, on='地级市层')
    city_merged_df['入样比例'] = city_merged_df['入样数'] / city_merged_df['总数']

    hangye_merged_df.to_excel('各行业入样情况统计.xlsx', index=False)
    city_merged_df.to_excel('各地级市入样情况统计.xlsx', index=False)
    print('各行业入样情况统计.xlsx已保存在当前目录下')
    print('各地级市入样情况统计.xlsx已保存在当前目录下')