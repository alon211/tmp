import win32com.client as win32
import sys
import numpy as np
import pandas as pd
sys.path.append('../')
import datetime
from usermodules import file_operate
from gernerate_recorder import recoder_process

###################################################
def read_SAK_file(path):
    validate=False
    value=None
    try:
        data=pd.read_csv(path,header=0,na_values=np.NaN,encoding='utf-8',dtype='str',keep_default_na=False)
    except Exception as e:
        return None
    check_h_orders_columns = np.array(['CompanyId', 'FactoryId', 'AreaId', 'OrderType', 'OrderName',
                              'OrderCode', 'ProductName', 'ProductCode', 'ProductModel',
                              'ProductStandardTime', 'JobTableName', 'PlanCount', 'CompletedCount',
                              'NgStepCount', 'WorkTime', 'Result', 'StartTime', 'EndTime', 'UserCode',
                              'UserName', 'Information', 'TorqueData', 'RotationTimeData',
                              'RotationAngleData', 'TraceData', 'ScrewSpecData', 'ProductVersion',
                              'BomSetName', 'BomSetCode', 'BomSetVersion', 'ProductInformation',
                              'CustomInformation', 'JobId', 'CommentData', 'PartsData'])
    check_h_job_steps_columns=np.array(['CompanyId', 'FactoryId', 'AreaId', 'ProductName', 'ProductCode',
       'ProductModel', 'JobId', 'JobTableName', 'StepNo', 'StepType', 'Result',
       'NgStepCount', 'StartTime', 'EndTime', 'WorkTime', 'StandardTime',
       'UserCode', 'UserName', 'Information', 'ProductVersion', 'BomSetName',
       'BomSetCode', 'BomSetVersion', 'CustomInformation', 'ScrewSpecData',
       'CommentData', 'TorqueData', 'RotationTimeData', 'RotationAngleData',
       'TraceData', 'PartsData'])
    filename=''
    if path.find('h_orders.csv')>-1:
        filename='h_orders.csv'
    if path.find('h_job_steps.csv')>-1:
        filename = 'h_job_steps.csv'
    if filename=='h_job_steps.csv':
        validate=(data.columns==check_h_job_steps_columns).all()
    elif filename=='h_orders.csv':
        validate = (data.columns == check_h_orders_columns).all()
        if validate:
            data = data[data['Result'] == '2']
    if validate:
        data['StartTime']=pd.to_datetime(data['StartTime'], format='%Y-%m-%d')
        data['EndTime']=pd.to_datetime(data['EndTime'], format='%Y-%m-%d')
        value=data
        return value
    else:
        return None
def get_Completed_Product_Data(ProductName:str,ProductVersion:str,TableName:str,StartDate:datetime.datetime,EndDate:datetime.datetime,data:pd.DataFrame):


    # 通过h_orders筛选出当日期范围内的所有完成订单，返回范围内的时间序列
    df=data[(data['ProductName']==ProductName)&
            # (data['AreaId']==TableName)&
            (data['ProductVersion']==ProductVersion)&
            (data['StartTime']>=StartDate)&
            (data['StartTime']<=EndDate)
    ]
    if df.empty:
        return None
    else:
        return df
def get_h_orders_complete_tracedata(h_orders_complete_data:pd.DataFrame):
    rst=None
def converter_data_to_RecodeData (data:pd.DataFrame,product_stepNo='1',
                                    cpu_stepNo='50',
                                    power_stepNo='42',
                                    ad_stepNo='18',
                                    driver_stepNO='15',
                                    LCD_stepNo='999',
                                    touchPanel_stepNo='990'
                                    )->recoder_process:
    '''
        cpu_stepNo等绑定实际的步号
    '''

    
    df=data
    '''
    把一个产品的所有条码记录分割开来 split做分割，explode负责将数组炸开
    源数据：
      1:31219031605170799007989190400006:::扫码工单号记录信息...
    变换后：
        0  1:31219031605170799007989190400006:::扫码工单号记录信息
        0                             15:18L6680:::扫码记录信息
        0                      18:V39AV191100158:::扫码记录信息
        0              42:31119031940602094:::取用电源模块并涂抹硅脂
        0                  50:AO2190321369000016:::扫码记录信息
    '''
    df=df.assign(TraceData=df['TraceData'].str.split(';')).explode('TraceData')
    '''
    新增2列将步骤号与条码分开来
    源数据：
        0  1:31219031605170799007989190400006:::扫码工单号记录信息
    变换后:
            stepNo                 codeNo
        0       1                   31219031605170799007989190400006

    '''
    df=df.assign(stepNo=(df['TraceData'].str.split(':',expand=True))[0],
                codeNo=(df['TraceData'].str.split(':',expand=True))[1])

    keys={product_stepNo:'产品序列号',#KEY值代表 步骤号码
            power_stepNo:'电源板',
            cpu_stepNo:'CPU板',
            LCD_stepNo:'液晶屏',
            ad_stepNo:'AD板',
            driver_stepNO:'驱动板',
            touchPanel_stepNo:'触摸屏'}
    keys2={'产品序列号':[0,3],  #keys2值代表模板表格内的输入的数据的索引号，gernerate_recorder.py的_init_data
                            '电源板':[2,5],
                            'CPU板':[3,5],
                            '液晶屏':[4,5],
                            'AD板':[5,5],
                            '驱动板':[6,5],
                            '触摸屏':[7,5]
                                    }
    '''
        增加一列步骤名称，通过stepNo查询Keys找到对应值，注意keys中的键不能重复
        变换后:
            stepNo   stepName              codeNo
        0       1     产品序列号              31219031605170799007989190400006

    '''
    df=df.assign(stepName=df['stepNo'].map(keys))
    '''
        增加一列数据索引（最终放入表格的数据索引号具体见类recoder_process中的init_data，
        通过stepName查询Keys2找到对应值，注意keys2中的键不能重复)
        变换后:
            stepNo   stepName              codeNo
        0       1     产品序列号              31219031605170799007989190400006
    '''
    df=df.assign(dataIndex=df['stepName'].map(keys2))
    
    df=df[['stepName','stepNo','codeNo','dataIndex','TraceData']]

    '''
        将stepName与codeNo列合并生成一个字典
        {stepName:[codeNo1，codeNo2]}
    '''
    rst=df.groupby('stepName').codeNo.apply(list).to_dict()
    rst=pd.DataFrame(rst)
# 创建标准空数据
    col_name=np.array(['产品序列号','电源板','CPU板','液晶屏','AD板','驱动板','触摸屏'])
    tmp=pd.DataFrame(columns=col_name,dtype=str)
# 将结果数据与空数据合并，outer是并集，inner是交集，因为可能会有一些字段没有
    # 行并集
    tmp1=pd.concat([tmp,rst],join='outer')

    pd.set_option('display.max_columns', None)#输出显示用
    # print(tmp1)
    rst=recoder_process()
    rst.all_data=[]

    df['tmp']=df.apply(lambda row:[row['dataIndex'],row['codeNo']],axis=1)
    df.groupby(df.index).tmp.apply(lambda row:write_data_to_recoderformat(row,rst)) 
    
    return rst
def write_data_to_recoderformat(input,rst:recoder_process):
    # 初始化currentdata
    rst.init_data()
    for item in input:
# 解压索引号和对应的数据 
        index,data=item
        tmp=rst.current_data[index[0]][index[1]]
        if tmp is None:
            rst.current_data[index[0]][index[1]]=data
        else:
            # 该条件主要针对产品序列号
            tmp=tmp.replace(" ","")
            rst.current_data[index[0]][index[1]]=tmp+data

    rst.add_current_data()
def get_h_steps_complete_tracedata(h_orders_complete_data:pd.DataFrame,h_steps_data:pd.DataFrame):
    # 从get_Completed_Product_Data获取到的所有完成产品数据然后查找h_job_steps里的对应的列数据
    '''

    :param h_orders_complete_data:
    :param h_steps_data:
    :return:
    '''
    rst=h_steps_data.drop(index=h_steps_data.index)
    for index,row in h_orders_complete_data.iterrows():
        tmp=h_steps_data[(h_steps_data['StartTime']>=row['StartTime'])&
                     (h_steps_data['EndTime']<=row['EndTime'])
        ]
        rst=pd.concat((rst,tmp))
    return rst
# AreaId ProductName ProductVersion StepNo StartTime EndTime
# 创建一个对话框：内容包含 作业台名称 作业程序版本 选择开始日期 结束日期 输出文档路径 输入文档路径(h_job_steps.csv h_order.csv)
# 先从h_order中确认完成产品的版本号，开始时间，结束时间


# data=pd.read_csv(r'E:\myproject\智能工作台\h_orders.csv',header=0,na_values='NULL',encoding='utf-8',dtype='str')
# check_columns=np.array(['CompanyId', 'FactoryId', 'AreaId', 'OrderType', 'OrderName',
#        'OrderCode', 'ProductName', 'ProductCode', 'ProductModel',
#        'ProductStandardTime', 'JobTableName', 'PlanCount', 'CompletedCount',
#        'NgStepCount', 'WorkTime', 'Result', 'StartTime', 'EndTime', 'UserCode',
#        'UserName', 'Information', 'TorqueData', 'RotationTimeData',
#        'RotationAngleData', 'TraceData', 'ScrewSpecData', 'ProductVersion',
#        'BomSetName', 'BomSetCode', 'BomSetVersion', 'ProductInformation',
#        'CustomInformation', 'JobId', 'CommentData', 'PartsData'])
# print(check_columns)
# print((data.columns==check_columns).all())
# #先筛选出完成的产品 result=2完成 =9未完成
# data=data[data['Result']=='2']
# # 先将作业台名称 作业程序版本 筛选出来
# df=data[(data['ProductName']=='四方监控屏111') & (data['ProductVersion']=='4') ]
# if df.empty:
#     print('完成产品中没有该产品版本号或者产品名称')
#     sys.exit(0)
# # 筛选出在日期范围内的数据索引号
# df_dt=pd.to_datetime(df['StartTime'],format='%Y-%m-%d')
# Select_StartTime=pd.to_datetime('2019-6-10')
# Select_EndTime=pd.to_datetime('2019-7-10')
# index=(df_dt[(Select_EndTime>df_dt) & (Select_StartTime<df_dt)].index)
# df=df.loc[index]
#
# data=pd.read_csv(r'E:\myproject\智能工作台\1.csv',header=0,na_values='NULL',encoding='utf-8',dtype='str')
# # 先将作业台名称 作业程序版本 筛选出来
# df=data[(data['ProductName']=='四方监控屏111') & (data['ProductVersion']=='4') ]
# # 筛选出在日期范围内的数据索引号
# df_dt=pd.to_datetime(df['StartTime'],format='%Y-%m-%d')
# Select_StartTime=pd.to_datetime('2019-6-10')
# Select_EndTime=pd.to_datetime('2019-7-10')
# index=(df_dt[(Select_EndTime>df_dt) & (Select_StartTime<df_dt)].index)
# df=df.loc[index]
#
# #筛选出产品序列号
# print(df[df['StepNo']=='1'].head(1)['TraceData'].iloc[0])
#
# print(df[df['BomSetCode']=='701050000477100'])

# data=read_SAK_file(r'C:\\Users\\MEACH\\project\\SAK\\h_job_steps.csv')
# print(data[data['TraceData']!='']['TraceData'].str.contains('3111809906'))