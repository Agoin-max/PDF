import pandas as pd
import json
import openpyxl
import sys
from Setting.Tools import FuTools
from Supplier_WenYe.Modify_Information import ModifyData
import time

class FormationDocuments:

    def Enter_Data(self, CompanyNumber):
        columns = json.loads(FuTools().open_json())[CompanyNumber]["Pandas_Columns"]
        name = json.loads(FuTools().open_json())[CompanyNumber]["CompanyName"]
        excel_path = json.loads(FuTools().open_json())["Path"]
        df = pd.read_excel("..\ExcelFile\PACKING.xlsx")
        ins = pd.read_excel("..\ExcelFile\INVOICE.xlsx")
        omg = pd.DataFrame(columns=columns)
        FormationController().enter_info(df, ins, name, omg)
        path = excel_path + time.strftime("%Y%m%d-%H%M%S") + "-" + name + "-入库计划单导入表.xlsx"
        omg.to_excel(path, index=False, startrow=1)
        df = ModifyData().invoke_pandas(path)
        df.to_excel(path, index=False, startrow=1)
        wb = openpyxl.load_workbook(path)
        sheet = wb['Sheet1']
        sheet.cell(1, 1).value = "入库计划单导入表"
        wb.save(path)
        wb.close()


class FormationController:

    def enter_info(self, df, ins, name, omg):
        for i in range(len(df)):
            data = {"商品类型": "",
                    "商品小类": "",
                    "品牌": df.iloc[i]["BRAND"],
                    "型号": df.iloc[i]["DESCRIPTION"],
                    "商品描述": "",
                    "产地": df.iloc[i]["ORIGIN"],
                    "单位": "",
                    "数量": str(df.iloc[i]["QTY(PCS)"]).replace(",", ""),
                    "报关单价": "",
                    # "件数": float(df.iloc[i]["C/NO"].split("~")[1]) - float(df.iloc[i]["C/NO"].split("~")[0]) + 1,
                    "件数": df.iloc[i]["C/NO"],
                    "净重": df.iloc[i]["N.W(KGS)"],
                    "毛重": df.iloc[i]["G.W(KGS)"],
                    "税号": "",
                    "SKU": "",
                    "供应商": name,
                    "期票天数": "",
                    "对应的采购": "",
                    "料号": df.iloc[i]["Cust.P/O#Header"],
                    "托盘数": "",
                    "箱号": "",
                    "备注": "",
                    "收款方": "",
                    "支付方式": "",
                    "关务品名": df.iloc[i]["Invo"]}
            for c in range(len(ins)):
                if df.iloc[i]["BRAND"] == ins.iloc[c]["BRAND"] and df.iloc[i]["DESCRIPTION"] == ins.iloc[c][
                    "DESCRIPTION"] and df.iloc[i]["Cust.P/O#Header"] == ins.iloc[c]["CTM"]:
                    customer = ins.iloc[c]["CUSTOMER"]
                    price = ins.iloc[c]["U/P"]
                    data["报关单价"] = price
                    data["对应的采购"] = customer
                    omg.loc[i + 1] = data
                    break



# FormationDocuments().Enter_Data("018129889")