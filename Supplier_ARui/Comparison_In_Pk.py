import pandas as pd
import sys


class ExcelComparison:

    def __init__(self):
        self.__excel_list = []
        self.__set_list = []

    def invoke_pandas(self, path):
        df = pd.read_excel(path)
        # print(df)
        df["QTY"] = df["QTY"].astype(str)
        for index, row in df.iterrows():
            append_dict = {}
            if "," in row["QTY"]:
                row["QTY"] = row["QTY"].replace(",", "")
            if "PCS" in row["QTY"]:
                row["QTY"] = row["QTY"].replace("PCS", "").strip()
            append_dict["Invo"] = row["Invo"]
            append_dict["PN"] = row["PN"]
            append_dict["QTY"] = row["QTY"]
            self.__excel_list.append(append_dict)
        self.formation_excel()
        df = pd.DataFrame(self.__set_list)
        # print(df)
        df.to_excel(path.rsplit(".", 1)[0] + "-比对.xlsx", index=False)

    def formation_excel(self):
        # 合并
        if len(self.__excel_list) > 1:
            for r in self.__excel_list:
                isExists = False
                if self.__set_list == []:
                    isExists = True
                    self.__set_list.append(r)
                else:
                    for c in self.__set_list:
                        # r["Invo"] == c["Invo"]
                        if r["PN"] == c["PN"]:
                            isExists = True
                            c["QTY"] = float(r["QTY"]) + float(c["QTY"])
                            # self.__set_list.append(r)
                if not isExists:
                    self.__set_list.append(r)
        else:
            self.__set_list = self.__excel_list

# ExcelComparison().invoke_pandas(r"C:\Users\windo\PycharmProjects\PDF\ExcelFile\PACKING.xlsx")
