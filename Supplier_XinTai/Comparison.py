import pandas as pd
from Supplier_XinTai.GetAExcel import FormationDocuments


class XinTaiComparison:

    def __init__(self, company_number=""):
        self.company_number = company_number

    def info_judge(self):
        df = pd.read_excel("..\ExcelFile\INVOICE-比对.xlsx",dtype=str)
        cn = pd.read_excel("..\ExcelFile\PACKING-比对.xlsx",dtype=str)
        Error_Info = []
        for index1, row1 in df.iterrows():
            isExist = False
            for index2, row2 in cn.iterrows():
                if row1["Invo"] == row2["Invo"] and row1["BRAND"] == row2["BRAND"] and row1["PART"] == \
                        row2["PART"] and row1["QTY"] == row2["QTY"]:
                    isExist = True
                    break
            if not isExist:
                Error_Info.append("Invo:{},品牌:{},型号:{},数量不一致！".format(row1["Invo"], row1["BRAND"], row1["PART"]))
                # return "品牌：{}，型号：{},数量不一致！".format(row1["BRAND"], row1["DESCRIPTION"])
        return Error_Info

    def formation_excel(self):
        if self.info_judge() == []:
            FormationDocuments().Enter_Data(self.company_number)
        else:
            return self.info_judge()


# print(XinTaiComparison().info_judge())
