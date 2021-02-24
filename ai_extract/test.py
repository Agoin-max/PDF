import pandas as pd
import openpyxl


# def tran_xlsx(self, fname):
#     excel = win32.gencache.EnsureDispatch('Excel.Application')
#     wb = excel.Workbooks.Open(fname)
#     wb.SaveAs(fname + "x", FileFormat=51)
#     wb.Close()
#     excel.Application.Quit()


# wb = openpyxl.load_workbook(r"C:\Users\86173\Desktop\创捷模板-委托进口确认单--双抬头 .xlsx")
# sheet = wb['报关']
# print(sheet.cell(4, 7).value)

a = "无牌/aaa"
b = a.find("无牌")
print(b)
