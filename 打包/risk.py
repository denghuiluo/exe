#!/usr/bin/env python
# -*- coding:utf-8 -*-

import pandas as pd
from tkinter import *
import tkinter.filedialog

# trans=pd.read_excel(r'C:\Users\Administrator\Desktop\数据处理\导出数据.xlsx',sheet_name='sheet1'
#                     ,usecols=['订单编号','司机id','乘客上车时间','到达目的地时间'])
# trans['差异时间']=pd.to_datetime(trans['到达目的地时间'])-pd.to_datetime(trans['乘客上车时间'])

# root=Tk()
# root.title('订单数据处理')
# root.geometry("400x300")
# def xz():
#     filename=tkinter.filedialog.askopenfilename()
#     if filename !='':
#         lb.config(text='您选择的文件是: '+filename)
#         trans = pd.read_excel(filename)
#         print(trans)
#     else:
#         lb.config(text='您没有选择任何文件')
#
# lb=Label(root,text='')
# lb.pack()
# btns=Button(root,text='选择源文件',background='red',command=xz)
# btns.place(x=100,y=100)
# # btns.pack()
# # btnd=Label(root,text='目标文件与源文件同目录')
# # btnd.pack()
# root.mainloop()




trans=pd.read_excel(r'C:\Users\Administrator\Desktop\数据处理\导出数据.xlsx')
trans['差异时间分']=(pd.to_datetime(trans['到达目的地时间'])-pd.to_datetime(trans['乘客上车时间'])).astype(str).apply(lambda x:x.split(':')[1])

def myfun1(x):
    count=0
    for value in x:
        if int(value)<=5:
            count+=1
    return count

df=trans.groupby('司机id').agg({
    '订单编号':'count',
    '差异时间分':myfun1
}).reset_index()

df['占比']=df['差异时间分']/df['订单编号']

resdf=df.rename(columns={'订单编号':'总订单','差异时间分':'短时间订单'})
resdf1=resdf[((resdf['总订单']>1)&(resdf['总订单']<=5)&(resdf['占比']>0.8))|(resdf['总订单']>5)&(resdf['占比']>0.6)]

resdf2=trans[trans['司机id'].isin(resdf1['司机id'].values.tolist())].sort_values(by=['司机id','下单时间'],ascending=False)

resdf2[['司机id','订单编号']]=resdf2[['司机id','订单编号',]].astype(str)
# resdf2['订单编号']=resdf2['订单编号'].astype(str)


resdf2.to_excel(r'C:\Users\Administrator\Desktop\测试.xlsx',index=False)



