def FillterStaff(path = r"C:\Users\liangbiguo\Desktop\结果", data = '201601-201705客户经理数据.xlsx'):
    # 设置路径
    os.chdir(path)
    # 加载数据
    df = pd.read_excel(data)
    # "今年和去年月均相差"和"201701-201705新增"的描述性统计
    ConstantNetGrowthDescribe = df['今年和去年月均相差'].describe()
    NewGrowthDescribe = df['201701-201705新增'].describe()
    # 默认是从低到高排序，所以75%是高排名分界点，25%是低 
    ConstantNetGrowthDescribe_3of4 = ConstantNetGrowthDescribe['75%']
    ConstantNetGrowthDescribe_1of4 = ConstantNetGrowthDescribe['25%']
    NewGrowthDescribe_3of4 = NewGrowthDescribe['75%']
    # 筛选两个都排名都在前25%的员工
    df0 = df[df['今年和去年月均相差'] >= ConstantNetGrowthDescribe_3of4]
    df1 = df0[df0['201701-201705新增'] >= NewGrowthDescribe_3of4]
    df1.to_excel('存量和新增均排名前25%的客户经理.xlsx', index = None)
    # 筛选较差的
    df2 = df[df['今年和去年月均相差'] <= ConstantNetGrowthDescribe_1of4]
    df3 = df2[df2['201701-201705新增'] == 0]
    df3.to_excel('存量排名后25%和新增为0的客户经理.xlsx', index = None)