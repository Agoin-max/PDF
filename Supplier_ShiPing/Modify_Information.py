import sys
import numpy as np
import pandas as pd


class ModifyData:

    def __init__(self):
        pass

    def invoke_pandas(self, path):
        df = pd.read_excel(path, header=1,dtype=str)
        df["件数"] = df["件数"].astype(str)
        df["净重"] = df["净重"].astype(str)
        for index, row in df[["净重", "毛重"]].iterrows():
            # print(float(row["净重"]))
            if row["净重"] == " " or row["净重"] == "nan" or float(row["净重"]) == 0.0:
                df.loc[index, "净重"] = np.NaN

        for index, row in df[["产地", "毛重"]].iterrows():
            if row["产地"] == "NAM":
                df.loc[index, "产地"] = "VIET NAM"

        for index, row in df[["件数", "净重", "毛重"]].iterrows():
            if row["净重"] != row["净重"]:
                df.loc[index, "件数"] = 0
            elif row["件数"].find("-") >= 0:
                df.loc[index, "件数"] = int(row["件数"].split("-")[1]) - int(
                    row["件数"].split("-")[0]) + 1
            else:
                df.loc[index, "件数"] = 1
        # print(df["件数"])
        # sys.exit()
        count = 0
        while count < 3:
            for index, row in df[["数量", "件数", "净重", "毛重"]].iterrows():
                quanty = 0
                if row["净重"] != row["净重"] and df.iloc[index - 1]["数量"] != 0:
                    quanty += int(df.iloc[index - 1]["数量"])
                    quanty += int(df.iloc[index]["数量"])
                    price_j = df.iloc[index - 1]["净重"]
                    price_m = df.iloc[index - 1]["毛重"]
                    for i in range(index + 1, len(df)):
                        if df.iloc[i]["净重"] != df.iloc[i]["净重"]:
                            quanty += int(df.iloc[i]["数量"])
                        else:
                            break
                    df.loc[index - 1, "净重"] = float(df.loc[index - 1, "数量"]) / float(quanty) * float(price_j)
                    df.loc[index - 1, "毛重"] = float(df.loc[index - 1, "数量"]) / float(quanty) * float(price_m)
                    df.loc[index, "净重"] = float(df.loc[index, "数量"]) / float(quanty) * float(price_j)
                    df.loc[index, "毛重"] = float(df.loc[index, "数量"]) / float(quanty) * float(price_m)
                    for i in range(index + 1, len(df)):
                        if df.iloc[i]["净重"] != df.iloc[i]["净重"]:
                            df.loc[i, "净重"] = float(df.loc[i, "数量"]) / float(quanty) * float(price_j)
                            df.loc[i, "毛重"] = float(df.loc[i, "数量"]) / float(quanty) * float(price_m)
                        else:
                            break
                    if df.loc[index, "净重"]:
                        break
            count += 1

        # print(df[["净重", "毛重"]])
        return df

# ModifyData().invoke_pandas()
