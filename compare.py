# -*- coding: utf-8 -*-
"""
Created on Sun Jun 26 10:51:56 2022

@author: Lenovo
"""

import xlwings as xw
import openpyxl
import pandas as pd
import xlrd
import time
from itertools import islice


file = "C:\\Users\\Lenovo\Desktop\\make money\\拆分\\test\\汇总2.xlsx"
sheet = "总表"

def xlwings_():
    wb = xw.Book(file)
    sht=wb.sheets[sheet]
    df = sht.range('A1').options(pd.DataFrame,header=1,index=False,expand='table').value
    print(df)

def open_pyxl():
    wb = openpyxl.load_workbook(file, data_only=True, read_only=True)
    ws = wb[sheet]
    data = ws.values
    cols = next(data)[0:]
    data = list(data)
    data = (islice(r, 0, None) for r in data)
    df = pd.DataFrame(data, columns=cols)
    wb.close()
    print(df)

def xlrd_():
    wb = xlrd.open_workbook(file)
    ws = wb.sheet_by_name(sheet) 
    data = [] #新建一个列表
    for r in range(ws.nrows): #将表中数据按行逐步添加到列表中，最后转换为list结构
        data1 = []
        for c in range(ws.ncols):
            data1.append(ws.cell_value(r,c))
        data.append(list(data1))
    cols = data[0]
    data = data[1:]
    df = pd.DataFrame(data, columns=cols)
    print(df)
'''
start = time.time()
xlwings_()
end = time.time()
print(end-start)

start = time.time()
open_pyxl()
end = time.time()
print(end-start)

start = time.time()
xlrd_()
end = time.time()
print(end-start)
'''
print(str('20211224') == '20211224')