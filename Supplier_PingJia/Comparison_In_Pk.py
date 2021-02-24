import pandas as pd
import sys


class ExcelComparison:

    def __init__(self):
        self.__excel_list = []
        self.__set_list = []

    def invoke_pandas(self, path):
        df = pd.read_excel(path)
        # print(df)
        df["QUANTITY"] = df["QUANTITY"].astype(str)
        for index, row in df.iterrows():
            append_dict = {}
            if "," in row["QUANTITY"]:
                row["QUANTITY"] = row["QUANTITY"].replace(",", "")
            append_dict["Invo"] = row["Invo"]
            append_dict["Item"] = row["Item"]
            # append_dict["PART"] = row["PART"]
            append_dict["QUANTITY"] = row["QUANTITY"]
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
                        if r["Invo"] == c["Invo"] and r["Item"] == c["Item"]:
                            isExists = True
                            c["QUANTITY"] = float(r["QUANTITY"]) + float(c["QUANTITY"])
                            # self.__set_list.append(r)
                if not isExists:
                    self.__set_list.append(r)
        else:
            self.__set_list = self.__excel_list


# ExcelComparison().invoke_pandas(r"C:\Users\windo\PycharmProjects\PDF\ExcelFile\PACKING.xlsx")
