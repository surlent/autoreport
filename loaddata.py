# -*- coding:utf8 -*-

import xlrd
import xlwt
import xlutils
import time


rawMonthSTR = input('Please Input Last Month :\n格式为20xx0x\n')  # 数据月份
rawMonthTS = time.strptime(rawMonthSTR,"%Y%m")  # 转换
# print(rawMonthTS)
fileName = r"E:/每月报表/" + rawMonthSTR + "test/" + rawMonthSTR + "销售清单.xlsx"  # test注意这是一个测试脚本
rawMaterailBook = xlrd.open_workbook(fileName)
sheet_name = rawMaterailBook.sheet_names()[0]
print(sheet_name)
sheet = rawMaterailBook.sheet_by_index(0)


class invoice(object):

    def __init__(self,number,custom,product,spec,weight,price,trademethed,packingfee):
        self.number = number
        self.custom = custom
        self.product = product
        self.spec = spec
        self.weight = weight
        self.price = price
        self.trademethed = trademethed
        self.packingfee = packingfee

invoiceList=()
# for rows in range(sheet.nrows):

row_data = sheet.row_values(0)  # 获得第1行的数据列表
print(row_data)
col_data = sheet.col_values(0)  # 获得第一列的数据列表，然后就可以迭代里面的数据了
print(col_data)
cell_value1 = sheet.cell_value(0, 0)  # 只有cell的值内容
print(cell_value1)
cell_value2 = sheet.cell(0, 0)  # 除了cell值内容外还有附加属性
print(cell_value2)
