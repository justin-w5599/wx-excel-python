# -*- encoding: utf-8 -*-
'''
@File    :   excelOprator.py
@Time    :   2020/05/27 15:50:49
@Author  :   Wang Junwen 
@Version :   1.0
@Contact :   junwen1938@163.com
'''
import os
import xlutils.copy  as copy
import xlrd,xlwt
import pandas as pd
import re
rec=re.compile(r'[\u4e00-\u9fa5]{2,3}')
redh=re.compile(r"1[\d]{10}")
os.chdir(r'C:\Users\Administrator\Desktop\2020审核认定材料\初访台帐认定、审核材料\初访通过21案中9案没材料的')
name_list=os.listdir()
filepath=r"C:\Users\Administrator\Desktop\台帐基础数据\县属各单位领导名单.xlsx"
filepath1=r"C:\Users\Administrator\Desktop\初访9案没有材料的.xlsx"
#读excel文件，worksheet:表名，header:列名,index_col:索引列,返回:DataFrame类型
def open_excel(filename=None,worksheet=0,header=0,index_col=0):
    ws=pd.read_excel(filename,sheet_name=worksheet,header=header,index_col=index_col,dtype=str)
    return ws
#批量写入工作表
def write_excel(df=pd.DataFrame(),filename=None,worksheet=0):
   with  pd.ExcelWriter(filename) as writer:
        df.to_excel(writer,sheet_name=worksheet)

#自动添写包保表
pathname=r"C:\Users\Administrator\Desktop\信访案件包保责任登记表.xls"
path_file=r"C:\Users\Administrator\Desktop\2020全台帐总表(攻坚、初访).xlsx"
path_save=r'C:\Users\Administrator\Desktop\2020审核认定材料\初访台帐认定、审核材料\初访通过21案中9案没材料的'
def write_preserve_sheet(pathname,listxfr):
    xl=xlrd.open_workbook(pathname,formatting_info=True)
    sheet=xl.sheet_by_index(0)
    # print(sheet.merged_cells,len(sheet.merged_cells))
    wb=copy.copy(wb=xl)
    style0 = xlwt.XFStyle()
    #设置字体
    font=xlwt.Font()
    font.name="仿宋_GB2312"
    font.height=12*20
    style0.font=font   
    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER
    alignment.wrap = 1
    style0.alignment=alignment
    #边框样式
    # borders = xlwt.Borders()
    # borders.left = xlwt.Borders.THICK # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUM_DASHED, THIN_DASH_DOTTED, MEDIUM_DASH_DOTTED, THIN_DASH_DOT_DOTTED, MEDIUM_DASH_DOT_DOTTED, SLANTED_MEDIUM_DASH_DOTTED, or 0x00 through 0x0D.
    # borders.right = xlwt.Borders.THICK
    # borders.top = xlwt.Borders.THICK
    # borders.bottom = xlwt.Borders.THICK
    # style0.borders=borders
    ws=wb.get_sheet(0)
    ws.write(6,2,listxfr[0],style0)
    ws.write(6,4,listxfr[1],style0)
    ws.write(6,6,listxfr[2],style0)
    ws.write(6,8,listxfr[3],style0)
    ws.write(7,2,listxfr[4],style0)
    ws.write(10,2,listxfr[5],style0)
    ws.write(10,4,listxfr[6],style0)
    ws.write(10,6,listxfr[7],style0)
    ws.write(10,8,listxfr[8],style0)
    ws.write(12,2,rec.match(listxfr[9]).group(),style0)
    ws.write(12,4,listxfr[10],style0)
    ws.write(12,8,redh.search(listxfr[9]).group(),style0)
    ws.write(14,2,rec.match(listxfr[12]).group(),style0)
    ws.write(14,4,listxfr[11],style0)
    ws.write(14,8,redh.search(listxfr[12]).group(),style0)
    ws.write(16,2,rec.match(listxfr[13]).group(),style0)
    ws.write(16,4,listxfr[11] + r"派出所",style0)
    ws.write(16,8,redh.search(listxfr[13]).group(),style0)

    ws.write(18,2,rec.match(listxfr[14]).group(),style0)
    ws.write(18,8,redh.search(listxfr[14]).group(),style0)
    savepath=path_save + "\\"  + listxfr[0] + "\\" + listxfr[0] + "信访案件包保责任登记表.xls"
    wb.save(savepath)
df1 = open_excel(filename=filepath,header=0) #县属各单位领导名单
df2 = open_excel(filename=filepath1,header=0) #信访人名单
# print(df2.head())
listxfr=[]
for l in name_list:
    # print(rec.match(l).group())
    listxfr=[]
    df2name=df2[df2["姓名"]==rec.match(l).group()]

    # print(df2name)
    df1name=df1[df1["姓名"]==df2name.iat[0,12]]
    # print(df1name)
    listxfr.append(df2name.iloc[0][0]) #00姓名
    listxfr.append(df2name.iloc[0][1]) #01id
    listxfr.append(df2name.iloc[0][2]) #02户籍
    listxfr.append(df2name.iloc[0][2]) #03住址
    listxfr.append(df2name.iloc[0][3]) #04诉求
    listxfr.append(df1name.iloc[0][1]) #05县领导姓名
    listxfr.append(df1name.iloc[0][0]) #06单位
    listxfr.append(df1name.iloc[0][2]) #07职务
    listxfr.append(df1name.iloc[0][3]) #08电话
    listxfr.append(df2name.iloc[0][7]) #09属事单位责任人姓名
    listxfr.append(df2name.iloc[0][6]) #10属事单位名称
    listxfr.append(df2name.iloc[0][8]) #11属地名称
    listxfr.append(df2name.iloc[0][9]) #12属地稳控责任人
    listxfr.append(df2name.iloc[0][11]) #13属地派出所
    listxfr.append(df2name.iloc[0][10]) #14村责任人
    write_preserve_sheet(pathname=pathname,listxfr=listxfr)
    # print(listxfr)
print("程序结束！")   
# print(type(listxfr),len(listxfr[0]))
# df=df1[df1["姓名"]=="高世忠"]
# print(df)
# print(type(df),df1.iloc[0,:][1])
# write_preserve_sheet(pathname)
