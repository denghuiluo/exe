import  PySimpleGUI as sg
import os
import pandas as pd
import openpyxl
from datetime import date, timedelta
sg.theme('LightBlue3')


yesterday = str((date.today() + timedelta(days=-1)).strftime("%#m.%#d"))
strchong = yesterday + '重庆呼我冲单奖'
strqu = yesterday + '重庆呼我时段奖'

def readdict(filepath):
    df = pd.DataFrame(pd.read_excel(filepath,'总调账版本'))
    df_rows = df.shape[0]
    wb=openpyxl.load_workbook(filepath)
    writer=pd.ExcelWriter(filepath,engine='openpyxl')
    writer.book=wb
    writer.sheets = dict((ws.title, ws) for ws in wb.worksheets)
    return df_rows,writer




layout=[
    [sg.Text('原始数据：',font=("微软雅黑",12)),sg.InputText('必填',key='file',size=(30,1),font=("微软雅黑",10),enable_events=True),sg.FileBrowse('打开',file_types=(("Text Files", "*.xls*"),),size=(10,1),font=("微软雅黑",10))],
    [sg.Text('司机数据：',font=("微软雅黑",12)),sg.InputText('必填',key='driverfile',size=(30,1),font=("微软雅黑",10)),sg.FileBrowse('打开',file_types=(("Text Files", "*.xls*"),),size=(10,1),font=("微软雅黑",10))],
    [sg.Text('连接字段：',font=("微软雅黑",12)),sg.Combo('',tooltip='选择连接的字段',font=("微软雅黑",10),default_value='',auto_size_text=True,size=(15,5),key='-keys-')],
    [sg.Text('冲单规则：', font=("微软雅黑", 12)),sg.InputText('可不填', key='chongfile', size=(30, 1), font=("微软雅黑", 10)),sg.FileBrowse('打开', file_types=(("Text Files", "*.xls*"),), size=(10, 1), font=("微软雅黑", 10))],
    [sg.Text('时段规则：', font=("微软雅黑", 12)),sg.InputText('可不填', key='qufile', size=(30, 1), font=("微软雅黑", 10)),sg.FileBrowse('打开', file_types=(("Text Files", "*.xls*"),), size=(10, 1), font=("微软雅黑", 10))],
    [sg.Text('目标目录：',font=("微软雅黑", 12)),sg.InputText('必填',key='folder',size=(30,1),font=("微软雅黑", 10),enable_events=True),sg.FolderBrowse('打开文件夹',size=(10, 1),font=("微软雅黑", 10))],
    [sg.Text('程序操作记录:',justification='center')],
    [sg.Output(size=(51,8),font=("微软雅黑",10)),
    sg.Button('开始执行',font=("微软雅黑",12),button_color='Orange'),
    sg.Button('关闭程序',font=("微软雅黑",12),button_color='red')]
]


#创建窗口
window=sg.Window('奖励计算工具',layout,font=("微软雅黑",12),default_element_size=(50,1),icon='chengzi.ico')

# 计算订单总数
def myfun(x):
    count = 0
    for value in x:
        hour = int(value.split()[1].split(':')[0])
        if (hour >= 9 and hour < 16):
            count += 1
    return count


#事件循环
while True:
    event,values=window.read()
    if event in (None,'关闭程序'):
        break
    if event=='file':
        filename = values['file']
        if os.path.exists(filename):
           orderdf=pd.read_excel(filename)
           keys=orderdf.columns.to_list()
           window["-keys-"].update(values=keys,font=("微软雅黑",10),size=(15,8))
    if event=='folder':
        if values['folder']:
            folder=values['folder']
            df1_rows, writerchong = readdict(folder+ "\\"+'重庆呼我冲单奖.xlsx')
            df2_rows, writerqu = readdict(folder+ "\\"+'重庆呼我时段奖.xlsx')
        else:
            df1_rows, writerchong = readdict(r'E:\chongqing\重庆呼我冲单奖.xlsx')
            df2_rows, writerqu = readdict(r'E:\chongqing\重庆呼我时段奖.xlsx')
    if event=='开始执行':
        #处理导入的订单数据
        df = orderdf.groupby('司机id').agg({
            '订单编号': 'count',
            '下单时间': myfun
        }).reset_index()
        df['司机id'] = df['司机id'].astype(str)
        df.rename(columns={'订单编号': '总订单', '下单时间': '区间订单'}, inplace=True)
        #处理司机数据
        driverfile=values['driverfile']
        driverdf = pd.read_excel(driverfile, converters={'司机id': str})
        #联合两表数据
        key = values['-keys-']
        total = df[df[key].isin(driverdf[key].values.tolist())]
        totaldf = pd.merge(total, driverdf, on=[key])
        resultdf = totaldf[['司机id', '姓名', '总订单', '区间订单']]

        #定义规则文件
        def guifile(code,file):
            chongdf = pd.read_excel(file)
            df = chongdf.sort_values(by='完单数', ascending=False)
            # 遍历规则中的数据组成规则字典函数
            dict = {}
            rows = df.shape[0]
            for i in range(rows):
                dict[df.iloc[i, 0]] = df.iloc[i, 1]
            dict[0] = 0
            for key, value in dict.items():
                if code >= int(key):
                    return int(value)
                else:
                    continue
        #读入冲单数据
        def chonggui(code):
            chongfile = values['chongfile']
            if os.path.exists(chongfile):
                value=guifile(code,chongfile)
                return value

        #读取区间数据
        def qugui(code):
            qufile = values['qufile']
            if os.path.exists(qufile):
                value = guifile(code, qufile)
                return value


        resultdf = resultdf.copy()
        resultdf['冲单金额'] = resultdf['总订单'].apply(chonggui)
        resultdf['时段金额'] = resultdf['区间订单'].apply(qugui)
        resultdf['冲单备注'] = strchong
        resultdf['冲单司机端'] = strchong
        resultdf['时段备注'] = strqu
        resultdf['时段司机端'] = strqu
        resultdf['司机id'] = resultdf['司机id'].astype(str)

        result1 = resultdf[['司机id', '姓名', '冲单金额', '冲单备注', '冲单司机端']]
        result1 = result1.copy()
        result1.rename(columns={'冲单金额': '金额', '冲单备注': '导入备注', '冲单司机端': '司机端说明'}, inplace=True)
        result11 = result1[result1['金额'] != 0]
        # result11.to_excel(r'E:\chongqing' + "\\" + strchong + ".xlsx",index=False)
        result11.to_excel(writerchong, sheet_name='总调账版本', startrow=df1_rows + 1, index=False, header=False)
        writerchong.save()
        writerchong.close()

        result2 = resultdf[['司机id', '姓名', '时段金额', '时段备注', '时段司机端']]
        resultdf = result2.copy()
        result2.rename(columns={'时段金额': '金额', '时段备注': '导入备注', '时段司机端': '司机端说明'}, inplace=True)
        result22 = result2[result2['金额'] != 0]
        # result22.to_excel(r'E:\chongqing' + "\\" + strqu + ".xlsx",index=False)
        result22.to_excel(writerqu, sheet_name='总调账版本', startrow=df2_rows + 1, index=False, header=False)
        writerqu.save()
        writerqu.close()
        print('----------已经完成----------\n')
