import sys

import pandas as pd
import json
import openpyxl
import time
from Setting.Tools import FuTools
from Supplier_FengYi.Modify_Information import ModifyData


class FormationDocuments:

    def Enter_Data(self, CompanyNumber):
        columns = json.loads(FuTools().open_json())[CompanyNumber]["Pandas_Columns"]
        name = json.loads(FuTools().open_json())[CompanyNumber]["CompanyName"]
        excel_path = json.loads(FuTools().open_json())["Path"]
        df = pd.read_excel("..\ExcelFile\PACKING.xlsx",dtype=str)
        ins = pd.read_excel("..\ExcelFile\INVOICE.xlsx",dtype=str)
        omg = pd.DataFrame(columns=columns)
        df["QTY"] = df["QTY"].astype(str)
        FormationController().enter_info(df, ins, name, omg)
        path = excel_path + time.strftime("%Y%m%d-%H%M%S") + "-" + name + "-入库计划单导入表.xlsx"
        omg.to_excel(path, index=False, startrow=1)
        # sys.exit()
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
                    "品牌": "",
                    "型号": df.iloc[i]["DESCRIPTION"],
                    "商品描述": "",
                    "产地": df.iloc[i]["C/O"],
                    "单位": "",
                    "数量": df.iloc[i]["QTY"].replace(",", ""),
                    "报关单价": "",
                    # "件数": float(df.iloc[i]["C/NO"].split("~")[1]) - float(df.iloc[i]["C/NO"].split("~")[0]) + 1,
                    "件数": df.iloc[i]["Carton"],
                    "净重": df.iloc[i]["NET"],
                    "毛重": df.iloc[i]["GROSS"],
                    "税号": "",
                    "SKU": "",
                    "供应商": name,
                    "期票天数": "",
                    "对应的采购": "",
                    "料号": "",
                    "托盘数": "",
                    "箱号": "",
                    "备注": "",
                    "收款方": "",
                    "支付方式": "",
                    "关务品名": df.iloc[i]["Invo"]}
            for c in range(len(ins)):
                if df.iloc[i]["Invo"] == ins.iloc[c]["Invo"] and df.iloc[i]["DESCRIPTION"] == ins.iloc[c][
                    "PN"]:
                    data["商品小类"] = ins.iloc[c]["Desc3"]
                    data["品牌"] = ins.iloc[c]["Brand"]
                    data["报关单价"] = ins.iloc[c]["PRICE"]
                    data["对应的采购"] = ins.iloc[c]["Cust"]
                    data["料号"] = ins.iloc[c]["PO"]
                    omg.loc[i + 1] = data
                    break
