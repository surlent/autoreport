import xlrd
import xlwt
import xlutils
import time


rawMonthSTR = "201706"  #  input('Please Input Last Month :\n格式为20xx0x\n')  # 数据月份
rawMonthTS = time.strptime(rawMonthSTR,"%Y%m")  # 转换
# print(rawMonthTS)
fileName = r"E:/每月报表/" + rawMonthSTR + "test/" + rawMonthSTR + "销售清单.xlsx"  # test注意这是一个测试脚本
rawMaterailBook = xlrd.open_workbook(fileName)
sheet_name = rawMaterailBook.sheet_names()[0]
#  print(sheet_name)
sheet = rawMaterailBook.sheet_by_index(0)


class invoice(object):

    def __init__(self,number,serialnumber,custom,product,weight,price,trademethod):
        self.number = number
        self.serialnumber = serialnumber
        self.custom = custom
        self.product = product
        self.weight = weight
        self.price = price
        self.trademethod = trademethod
        #  self.packingfee = packingfee


print(sheet.cell(2, 4).value)

invoiceList=[]
for row in range(1,4):
    invoiceList[row]= invoice(
    sheet.cell(row, 0).value,
    sheet.cell(row, 1).value,
    sheet.cell(row, 3).value,
    sheet.cell(row, 4).value,
    sheet.cell(row, 5).value,
    sheet.cell(row, 6).value,
    sheet.cell(row, 2).value
    #  sheet.cell(row, 10).value,
    #  sheet.cell(row, 8).value
    )

print(invoiceList)

