import pandas as pd


class ExcelInfo:

    def invoke_pandas(self):
        df = pd.read_excel("../ExcelFile/HongYi.xlsx")
        # print(df)
        C_NO = []
        Brand = []
        Description_1 = []
        Description_2 = []
        Customer = []
        QTY = []
        NW = []
        GW = []
        df["Brand"] = df["Brand"].astype(str)
        df["QTY"] = df["QTY"].astype(str)
        df["NW"] = df["NW"].astype(str)
        df["GW"] = df["GW"].astype(str)
        for index, row in df.iterrows():
            if row["C/NO"] == "Total:":
                break
            if row["C/NO"] == row["C/NO"]:
                C_NO.append(row["C/NO"])
            if row["Brand"].find("cm") < 0 and row["Brand"] != "nan":
                Brand.append(row["Brand"])
                Description_1.append(row["Description_c"])
            if row["Brand"].find("cm") >= 0:
                Description_2.append(row["Description_c"])
            if row["Customer PO"] == row["Customer PO"] and row["Customer PO"] != " ":
                Customer.append(row["Customer PO"])
            if row["QTY"].find("PCS") >= 0:
                QTY.append(row["QTY"].replace("PCS", "").strip())
                if row["NW"] == "nan":
                    NW.append("")
                else:
                    NW.append(row["NW"].replace("kg", ""))
                if row["GW"] == "nan":
                    GW.append("")
                else:
                    GW.append(row["GW"].replace("kg", ""))

        dict_values = {}
        dict_values["C/NO"] = C_NO
        dict_values["Brand"] = Brand
        dict_values["Description_1"] = Description_1
        dict_values["Description_2 "] = Description_2
        dict_values["Customer"] = Customer
        dict_values["QTY"] = QTY
        dict_values["NW"] = NW
        dict_values["GW"] = GW

        ins = pd.DataFrame(dict_values)
        ins.to_excel("../ExcelFile/PACKING.xlsx", index=False)


ExcelInfo().invoke_pandas()
