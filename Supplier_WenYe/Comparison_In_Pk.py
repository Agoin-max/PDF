import pandas as pd


class ExcelComparison:

    def __init__(self):
        self.__excel_list = []
        self.__set_list = []

    def invoke_pandas(self, path):
        df = pd.read_excel(path)
        df["QTY(PCS)"] = df["QTY(PCS)"].astype(str)
        for index, row in df.iterrows():
            append_dict = {}
            if "," in row["QTY(PCS)"]:
                row["QTY(PCS)"] = row["QTY(PCS)"].replace(",", "")
            append_dict["DESCRIPTION"] = row["DESCRIPTION"]
            append_dict["BRAND"] = row["BRAND"]
            append_dict["QTY(PCS)"] = row["QTY(PCS)"]
            append_dict["Invo"] = row["Invo"]
            self.__excel_list.append(append_dict)
        self.formation_excel()
        df = pd.DataFrame(self.__set_list)
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
                        if r["DESCRIPTION"] == c["DESCRIPTION"] and r["BRAND"] == c["BRAND"] and r["Invo"] == c["Invo"]:
                            isExists = True
                            c["QTY(PCS)"] = float(r["QTY(PCS)"]) + float(c["QTY(PCS)"])
                            # self.__set_list.append(r)
                if not isExists:
                    self.__set_list.append(r)
        else:
            self.__set_list = self.__excel_list

# ExcelComparison().invoke_pandas(r"C:\Users\windo\PycharmProjects\PDF\ExcelFile\PACKING.xlsx")
