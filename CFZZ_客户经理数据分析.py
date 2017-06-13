# load packages and set workpath
import numpy as np
import pandas as pd
import os
import warnings
warnings.filterwarnings("ignore")

# 读取原始数据名称
def GetRawDataName(path = r'E:\Work\CFZZ\Data\RawData'):
    os.chdir(path)
    name = os.listdir()
    return name

# 建立预处理后的数据名称，需要传入原始数据名称
# 原始数据为分页的xls文件，处理后数据名称为可容纳大数据的csv格式
def RowToPreName(RawDataName, WitchFormat = '.csv'):
    name = []
    for i in RawDataName:
        ss = i.split('.')[0] + WitchFormat
        name.append(ss)
    return name

# 需要手工查看每个xls文件具有多少个分页，然后把数字放在一个列表中
def SheetOfExcel():
    n1 = 3 #2016年01-2016年06.xls 有30个分页
    n2 = 2 #2016年07-2016年12.xls 有29个分页
    n3 = 2 #201701-201705存量.xls 有21个分页
    n4 = 2 #201701-201705新增.xls 有2个分页
    n_sheet = [n1, n2 ,n3, n4]
    return n_sheet

# 把全部分页Excel合并为单个CSV文件
def ConcatToCsv():
    # 加载原始数据名称
    RawDataName = GetRawDataName()
    # 加载预处理数据名称
    PreDataName = RowToPreName(RawDataName)
    # 加载每个文件的sheet分页数量
    n_sheet = SheetOfExcel()
    
    # 开始合成四个文件
    for i in range(len(RawDataName)):
        os.chdir(r'E:\Work\CFZZ\Data\RawData')
        df = pd.read_excel(RawDataName[i], sheetname = 0)
        for j in range(1, n_sheet[i]):
            df1 = pd.read_excel(RawDataName[i], sheetname = j)
            df = df.append(df1,ignore_index = True)
        os.chdir(r'E:\Work\CFZZ\Data\PreData')
        df.to_csv(PreDataName[i], index = None)

# 把16年的连个CSV合并
# 合并了后编码又出了问题，而且因为数据太大，无法用Notepad更改编码，因此就不合并了
def Concat2Csv():
    # 加载原始数据名称
    RawDataName = GetRawDataName()
    # 加载预处理数据名称
    PreDataName = RowToPreName(RawDataName)
    os.chdir(r'E:\Work\CFZZ\Data\PreData2')
    df = pd.read_csv(PreDataName[0])
    df1 = pd.read_csv(PreDataName[1])
    df = df.append(df1,ignore_index = True)
    os.chdir(r'E:\Work\CFZZ\Data\PreData3')
    df.to_csv('2016年01-2016年12.csv', index = None)
    
# 数据预处理部分到此截止了
###################################################################################################

# 计算单个员工某一时段的客户总交易量
def CalculateSumToEachPerson():
    # 加载文件名
    RawDataName = GetRawDataName(path = r'E:\Work\CFZZ\Data\PreData2')
    # 设置需要保存的文件名
    PreDataName = RowToPreName(RawDataName, '.xlsx')
    # 计算成交总和、合并单个员工、匹配营业部等
    for i in range(len(RawDataName) - 1):
        os.chdir("E:\Work\CFZZ\Data\PreData2")
        # 逐个加载数据
        df = pd.read_csv(RawDataName[i])
        # 计算总的成交金额
        df['总成交金额'] = df.ix[:, 8] + df.ix[:, 9] + df.ix[:, 10] + df.ix[:, 11] + df.ix[:, 12] + df.ix[:, 13]
        # 删掉没用的行
        df = df.drop(df.ix[:,[0, 1, 2, 4, 5, 8, 9, 10, 11, 12, 13]], axis=1)
        # 得到每个员工的客户一年交易总金额，设为pd表
        df1 = df.groupby('人员编号').agg(np.sum).reset_index()
        # 删掉原表的最后一列
        df = df.drop(df.ix[:,[-1]], axis=1)
        # df表去重，方便后面的合并
        df.drop_duplicates(['人员编号'],keep='last',inplace=True)
        # 合并，形成新的表
        df1 = pd.merge(df1,df,how='left',left_on='人员编号',right_on='人员编号')
        os.chdir("E:\Work\CFZZ\Data\WorkData")
        df1.to_excel(PreDataName[i], index = None)

# 把单个员工的几个数据表合并起来
def ConcatEachPersonData():
    # 加载文件名
    RawDataName = GetRawDataName(path = r'E:\Work\CFZZ\Data\WorkData')
    # 设定列名
    PreDataName = RowToPreName(RawDataName, '')
    # 加载第一个数据，并且更改列名
    df = pd.read_excel(RawDataName[0])
    df = df.rename(columns = {'总成交金额': PreDataName[0]})
    # 这一块代码存在bug， 下面括号里填2没错，但是填4就错了，分开步骤也没问题
    for i in range(1, 2):
        # 依次加载数据，并且更改列名
        df1 = pd.read_excel(RawDataName[i])
        df1 = df1.rename(columns = {'总成交金额': PreDataName[i]})
        # 合并数据
        df = pd.merge(df1, df, how='outer', on='人员编号')
        # 把人员名称用Y填充进来，保证最大的填充
        for j in range(len(df)):
            if df['人员名称_x'][j] is np.nan:
                df['人员名称_x'][j] = df['人员名称_y'][j]
        for j in range(len(df)):
            if df['经营单元_x'][j] is np.nan:
                df['经营单元_x'][j] = df['经营单元_y'][j]
        # 删掉没用的行
        df = df.drop(df.ix[:,[-1, -2]], axis=1)
        # 把名字改回去，方便下一次合并
        df = df.rename(columns = {'经营单元_x' : '经营单元', '人员名称_x':'人员名称'})
    os.chdir("E:\Work\CFZZ\Data\WorkData1")
    df.to_excel("201601-201705.xlsx", index = None)