# -*- encoding: utf-8 -*-
'''
@File    :   cp_file.py
@Time    :   2020/06/19 10:01:50
@Author  :   Wang Junwen 
@Version :   1.0
@Contact :   junwen1938@163.com
'''
import os
import pandas as pd

os.chdir('C:\\Users\\Administrator\\Desktop\\2020审核认定材料\\初访台帐认定、审核材料\\20200617初访台账剩余37案整理\\初访27案有答复意见')
fpath='C:\\Users\\Administrator\\Desktop\\答复意见书.docx'
fpath1='C:\\Users\\Administrator\\Desktop\\2020审核认定材料\\初访台帐认定、审核材料\\20200617初访台账剩余37案整理\\初访27案有答复意见\\' 
name_list=os.listdir()
for n in name_list:
    file1='copy ' + fpath + ' ' + fpath1 + n + "\\" + n +r"答复意见书" + r".docx"
    os.system(file1)
   