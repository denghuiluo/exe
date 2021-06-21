#!/usr/bin/env python
# -*- coding:utf-8 -*-

import pandas as pd
from datetime import date, timedelta
import openpyxl


def readdict(filepath):
    df = pd.DataFrame(pd.read_excel(filepath,sheet_name='总调账版本'))
    df_rows = df.shape[0]
    wb=openpyxl.load_workbook(filepath)
    writer=pd.ExcelWriter(filepath,engine='openpyxl')
    writer.book=wb
    writer.sheets = dict((ws.title, ws) for ws in wb.worksheets)
    return df_rows,writer

df1_rows,writerchong=readdict('E:\chongqing\重庆呼我冲单奖.xlsx')
df2_rows,writerqu=readdict('E:\chongqing\重庆呼我时段奖.xlsx')


yesterday = str((date.today() + timedelta(days=-1)).strftime("%#m.%#d"))
strchong = yesterday + '重庆呼我冲单奖'
strqu = yesterday + '重庆呼我时段奖'

# 挑选呼我司机
# cdf = pd.read_excel(r'E:\chongqing\重庆司机.xlsx')
# cdf1 = cdf[cdf['司机区分'] == '呼我']
# driverdf = cdf1[['司机id', '姓名']]


driverdf = pd.read_excel(r'E:\chongqing\重庆司机.xlsx',converters={'司机id':str})
# driverdf['司机id']=driverdf['司机id'].astype(str)
# print(driverdf)



# 计算订单总数
def myfun(x):
    count = 0
    for value in x:
        hour = int(value.split()[1].split(':')[0])
        if (hour >= 17 and hour < 19):
            count += 1
    return count


orderdf = pd.read_excel(r'E:\chongqing\导出数据.xlsx')
# orderdf['下单时间']=pd.to_datetime(orderdf['下单时间'])
df = orderdf.groupby('司机id').agg({
    '订单编号': 'count',
    '下单时间': myfun
}).reset_index()

df['司机id']=df['司机id'].astype(str)
df.rename(columns={'订单编号': '总订单', '下单时间': '区间订单'}, inplace=True)
# df.to_excel(r'E:\chongqing' + "\\"  + "1.xlsx",index=False)

total = df[df['司机id'].isin(driverdf['司机id'].values.tolist())]

totaldf = pd.merge(total, driverdf, on=['司机id'])

resultdf = totaldf[['司机id', '姓名', '总订单', '区间订单']]


# 方案一：
# def chongdan(i):
#     if i >= 26:
#         return 80
#     elif i >= 20:
#         return 55
#     elif i >= 15:
#         return 32
#     elif i >= 10:
#         return 18
#     else:
#         return 0


def chongdan(i):
    if i >= 29:
        return 120
    elif i >= 21:
        return 70
    elif i >= 15:
        return 40
    elif i >= 8:
        return 15
    else:
        return 0



def shiduan(i):
    if i >= 6:
        return 15
    elif i >= 4:
        return 8
    else:
        return 0

resultdf=resultdf.copy()
resultdf['冲单金额'] = resultdf['总订单'].apply(chongdan)
resultdf['时段金额'] = resultdf['区间订单'].apply(shiduan)
resultdf['冲单备注'] = strchong
resultdf['冲单司机端'] = strchong
resultdf['时段备注'] = strqu
resultdf['时段司机端'] = strqu
resultdf['司机id'] = resultdf['司机id'].astype(str)

result1 = resultdf[['司机id', '姓名', '冲单金额', '冲单备注', '冲单司机端']]
result1=result1.copy()
result1.rename(columns={'冲单金额': '金额', '冲单备注': '导入备注', '冲单司机端': '司机端说明'}, inplace=True)
result11 = result1[result1['金额'] != 0]
result11.to_excel(r'E:\chongqing' + "\\" + strchong + ".xlsx",index=False)
# result11.to_excel(writerchong,sheet_name='总调账版本',startrow=df1_rows+1,index=False,header=False)
# writerchong.save()
# writerchong.close()

result2 = resultdf[['司机id', '姓名', '时段金额', '时段备注', '时段司机端']]
result2=result2.copy()
result2.rename(columns={'时段金额': '金额', '时段备注': '导入备注', '时段司机端': '司机端说明'}, inplace=True)
result22 = result2[result2['金额'] != 0]
result22.to_excel(r'E:\chongqing' + "\\" + strqu + ".xlsx",index=False)
# result22.to_excel(writerqu,sheet_name='总调账版本',startrow=df2_rows+1,index=False,header=False)
# writerqu.save()
# writerqu.close()