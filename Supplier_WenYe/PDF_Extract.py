import sys
from decimal import Decimal
import pandas as pd
import json
from Setting.Father import DataPdf
from Setting.Tools import FuTools
from Supplier_WenYe.Comparison_In_Pk import ExcelComparison
import re


class ExtractPdf(DataPdf):

    def __init__(self, pdf_pages=0, pdf_obj=object, company_number=""):
        self.pdf_pages = pdf_pages
        self.pdf_obj = pdf_obj
        self.company_number = company_number

    def get_info(self):
        """文曄,公司编码:018129889"""
        # dict_values = {}
        meau = ""
        config_info = json.loads(FuTools().open_json())
        Packing_Df = pd.DataFrame()
        Invo_list = []
        Packing_Df.iterrows()
        for r in range(self.pdf_pages):
            dict_values = {}
            meau = ""
            dict_words = self.pdf_obj.pages[r].extract_words()
            start = 0
            end = 0
            Invo = ""
            try:
                pattern = re.compile(r"INVOICE NO: (.*?)\n", re.S)
                Invo = re.findall(pattern, self.pdf_obj.pages[r].extract_text())[0]
            except:
                print("未提取Invo成功！")

            # print(dict_words)
            for item in dict_words:
                for c in config_info[self.company_number]["Meau"]:
                    if item["text"] == c:
                        meau = self.get_invoice(dict_values, dict_words, end, config_info, c, Invo)
                        break
                if meau:
                    break

            onepage_df = pd.DataFrame()
            for title in dict_values:
                onetitle_df = pd.DataFrame(dict_values.get(title), columns=[title])
                onepage_df = pd.concat([onepage_df, onetitle_df], axis=1)
            Packing_Df = Packing_Df.append(onepage_df)

            for i in range(len(onepage_df)):
                Invo_list.append(Invo)


        Packing_Df["Invo"] = Invo_list
        Packing_Df.to_excel("../ExcelFile/" + meau + ".xlsx", index=False)
        path = "../ExcelFile/" + meau + ".xlsx"
        # sys.exit()
        # path = self.formate_excel(dict_values, meau)
        ExcelComparison().invoke_pandas(path)


    def get_invoice(self, dict_values, dict_words, end, config_info, c, Invo):
        for item in dict_words:
            if item["text"] == config_info[self.company_number]["Meau"][c]["end_first"] or item["text"] == \
                    config_info[self.company_number]["Meau"][c]["end_second"]:
                end = dict_words.index(item)
                break
        list_columns = config_info[self.company_number]["Meau"][c]['columns']
        for r in list_columns:
            for item in dict_words:
                if item["text"] == r:
                    start = dict_words.index(item)
                    # print(start)
                    x0 = dict_words[start]["x0"]
                    x1 = dict_words[start]["x1"]
                    top = dict_words[start]["top"]
                    bottom = dict_words[start]["bottom"]
                    Controller().get_info(start, end, dict_words, x0, top, bottom, x1, dict_values, r, config_info,
                                          self.company_number, c)
                    break
        return c

    # def formate_excel(self, dict_values, meau):
    #     df = pd.DataFrame(dict_values)
    #     df.to_excel("../ExcelFile/" + meau + ".xlsx", index=False)
    #     return "../ExcelFile/" + meau + ".xlsx"


class Controller:

    def get_info(self, start, end, dict_words, x0, top, bottom, x1, dict_values, r, config_info, company_number, index):
        for i in range(start, end):
            for c in range(1, 100):
                x0_left = config_info[company_number]["Meau"][index]["x0_left"]
                x1_right = config_info[company_number]["Meau"][index]["x1_right"]
                move = config_info[company_number]["Meau"][index]["move"]
                top_0 = config_info[company_number]["Meau"][index]["top_0"]
                bottom_0 = config_info[company_number]["Meau"][index]["bottom_0"]
                if x0 - Decimal(x0_left) <= dict_words[i]["x0"] <= x0 + Decimal(x0_left) and \
                        top + Decimal(move) * c - Decimal(top_0) <= dict_words[i]["top"] <= top + Decimal(
                    move) * c + Decimal(top_0) and \
                        bottom + Decimal(move) * c - Decimal(bottom_0) <= dict_words[i]["bottom"] <= bottom + Decimal(
                    move) * c + Decimal(bottom_0) or \
                        x1 - Decimal(x1_right) <= dict_words[i]["x1"] <= x1 + Decimal(x1_right) and \
                        top + Decimal(move) * c - Decimal(top_0) <= dict_words[i]["top"] <= top + Decimal(
                    move) * c + Decimal(top_0) \
                        and bottom + Decimal(move) * c - Decimal(bottom_0) <= dict_words[i][
                    "bottom"] <= bottom + Decimal(move) * c + Decimal(bottom_0):
                    try:
                        if r == "ORIGIN":
                            try:
                                if int(dict_words[i + 1]["text"]):
                                    dict_values.setdefault(r, []).append(dict_words[i]["text"])
                            except:
                                dict_values.setdefault(r, []).append(
                                    dict_words[i]["text"] + "," + dict_words[i + 1]["text"])
                        else:
                            dict_values.setdefault(r, []).append(dict_words[i]["text"])
                    except:
                        dict_values.setdefault(r, []).append("")
                    if not dict_words[i] and i > 50:
                        break
