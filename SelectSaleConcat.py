import numpy as np
import pandas as pd
import os

# 读取原始数据名称
def GetDataName(ReadPath):
    os.chdir(ReadPath)
    name = os.listdir()
    return name[:4]

# 建立预处理后的数据名称，需要传入原始数据名称
def ChangeDataName(DataName, WhichFormat = '.xlsx'):
    name = []
    for i in DataName:
        ss = i.split('.')[0] + WhichFormat
        name.append(ss)
    return name

##########################分类求和、合并数据#############################

#按“客户编号”列去重，方便后面的分组求和后的字符匹配
def  DeDuplication(df, ConcatID = '客户编号'):
    # 去重，为了后面匹配汉子字符
    df1 = df.drop_duplicates([ConcatID])
    # 把需要求和的、客户编号、时间删掉
    df1 = df1.drop(df1.ix[ : ,[-1,-2,-3,-4,-5,-6]], axis = 1)
    # 重新设置为文档
    df1 = df1.reset_index()
    # 把第一列删掉
    df1 = df1.drop(df1.ix[ : ,[0]], axis = 1)
    return df1

# 计算客户周期内的各种数据之和
def SumClientData(df, ConcatID = '客户编号'):
    df1 = df.groupby(ConcatID).agg(np.sum).reset_index()
    df1 = df1.drop(df1.ix[ : , ['月份','人员编号']], axis = 1)
    return df1

# 进行匹配并保存
def MatchClientData(df1, df2):
    df3 = pd.merge(df1, df2, how = 'left', on='客户编号')
    return df3

# 上面的几个步骤合并
def SumToMatch(df, ConcatID = '客户编号'):
    df1 = DeDuplication(df = df, ConcatID = ConcatID)
    df2 = SumClientData(df = df, ConcatID = ConcatID)
    df3 = MatchClientData(df1 = df1, df2 = df2)
    return df3
    
#####################筛选营业部+分类求和合并#########################
def SelcetSaleAndSumMatch(df, SaleName = '天津烟台路(单元)', ConcatID = '客户编号'):
    # 筛选营业部
    df1 = df[(df['经营单元'] == SaleName)]
    df2 = SumToMatch(df = df1, ConcatID = ConcatID)
    return df2

#############筛选营业部+分类求和+2016年2个表合并+再次求和############
def SlectSumConcatSum(df_1, df_2, SaleName = '天津烟台路(单元)', ConcatID = '客户编号'):
    df1 = SelcetSaleAndSumMatch(df_1, SaleName = SaleName, ConcatID = ConcatID)
    df2 = SelcetSaleAndSumMatch(df_2, SaleName = SaleName, ConcatID = ConcatID)
    df3 = df1.append(df2)
    df4 = SumToMatch(df = df3, ConcatID = '客户编号')
    return df4
    
###########################筛选2016数据并合并#######################
ReadPath = r'E:\Work\CFZZ\Data\PreData2'
SavePath = r'E:\Work\CFZZ\Data\WorkData3\Step1'
DataName = GetDataName(ReadPath = ReadPath)
df_1 = pd.read_csv(DataName[0])
df_2 = pd.read_csv(DataName[1])
Name = ['天津烟台路(单元)', '长沙八一路(单元)', '长沙芙蓉中路(单元)', '深圳深南大道(单元)', '长沙总部营业部(单元)',
       '东莞黄金路证券营业部(单元)', '邵东金龙大道(单元)', '浏阳劳动中路(单元)', '长沙星沙北路(单元)', '长沙万芙路(单元)',
       '三门上洋路（单元）', '北京三环中路(单元)', '太原营业部(单元)', '沈阳北陵大街(单元)', '合肥营业部(单元)',
       '南昌营业部(单元)', '青岛山东路(单元)', '郑州营业部(单元)', '成都营业部(单元)', '贵阳营业部(单元)',
       '昆明营业部(单元)', '西安大庆路(单元)', '兰州营业部(单元)', '广州营业部(单元)', '中山营业部(单元)',
       '重庆营业部(单元)', '石家庄营业部(单元)', '哈尔滨营业部(单元)', '福州营业部(单元)']

for i in range(len(Name)):
    df = SlectSumConcatSum(df_1 = df_1, df_2 = df_2, SaleName = Name[i], ConcatID = '客户编号')
    name = Name[i] + '_2016.xlsx'
    os.chdir(SavePath)
    df.to_excel(name, index = None)


######################筛选201701-201705存量数据###################
ReadPath = r'E:\Work\CFZZ\Data\PreData2'
SavePath = r'E:\Work\CFZZ\Data\WorkData3\Step2'
DataName = GetDataName(ReadPath = ReadPath)
df = pd.read_csv(DataName[2])
Name = ['天津烟台路(单元)', '长沙八一路(单元)', '长沙芙蓉中路(单元)', '深圳深南大道(单元)', '长沙总部营业部(单元)',
       '东莞黄金路证券营业部(单元)', '邵东金龙大道(单元)', '浏阳劳动中路(单元)', '长沙星沙北路(单元)', '长沙万芙路(单元)',
       '三门上洋路（单元）', '北京三环中路(单元)', '太原营业部(单元)', '沈阳北陵大街(单元)', '合肥营业部(单元)',
       '南昌营业部(单元)', '青岛山东路(单元)', '郑州营业部(单元)', '成都营业部(单元)', '贵阳营业部(单元)',
       '昆明营业部(单元)', '西安大庆路(单元)', '兰州营业部(单元)', '广州营业部(单元)', '中山营业部(单元)',
       '重庆营业部(单元)', '石家庄营业部(单元)', '哈尔滨营业部(单元)', '福州营业部(单元)']

for i in range(len(Name)):
    df1 = SelcetSaleAndSumMatch(df = df, SaleName = Name[i], ConcatID = '客户编号')
    name = Name[i] + '_2017存量.xlsx'
    os.chdir(SavePath)
    df1.to_excel(name, index = None)


######################筛选201701-201705新增数据###################
ReadPath = r'E:\Work\CFZZ\Data\PreData2'
SavePath = r'E:\Work\CFZZ\Data\WorkData3\Step3_2017新增'
DataName = GetDataName(ReadPath = ReadPath)
df = pd.read_csv(DataName[3])
Name = ['天津烟台路(单元)', '长沙八一路(单元)', '长沙芙蓉中路(单元)', '深圳深南大道(单元)', '长沙总部营业部(单元)',
       '东莞黄金路证券营业部(单元)', '邵东金龙大道(单元)', '浏阳劳动中路(单元)', '长沙星沙北路(单元)', '长沙万芙路(单元)',
       '三门上洋路（单元）', '北京三环中路(单元)', '太原营业部(单元)', '沈阳北陵大街(单元)', '合肥营业部(单元)',
       '南昌营业部(单元)', '青岛山东路(单元)', '郑州营业部(单元)', '成都营业部(单元)', '贵阳营业部(单元)',
       '昆明营业部(单元)', '西安大庆路(单元)', '兰州营业部(单元)', '广州营业部(单元)', '中山营业部(单元)',
       '重庆营业部(单元)', '石家庄营业部(单元)', '哈尔滨营业部(单元)', '福州营业部(单元)']

for i in range(len(Name)):
    df1 = SelcetSaleAndSumMatch(df = df, SaleName = Name[i], ConcatID = '客户编号')
    name = Name[i] + '_2017新增.xlsx'
    os.chdir(SavePath)
    df1.to_excel(name, index = None)
