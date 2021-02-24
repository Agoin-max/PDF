from decimal import Decimal
import pandas as pd
import json
from Setting.Father import DataPdf
from Setting.Tools import FuTools
import sys
import re

from Supplier_PingJia.Comparison_In_Pk import ExcelComparison


class ExtractPingJiaWorld(DataPdf):

    def __init__(self, pdf_pages=0, pdf_obj=object, company_number=""):
        self.pdf_pages = pdf_pages
        self.pdf_obj = pdf_obj
        self.company_number = company_number

    def get_info(self):
        # dict_values = {}
        meau = ""
        c = ""

        config_info = json.loads(FuTools().open_json())
        Invoice_Df = pd.DataFrame()
        Packing_Df = pd.DataFrame()
        coo_list = []
        Invo_list = []
        Invo_listone = []
        Invo_coo_list = []
        for r in range(self.pdf_pages):
            dict_words = self.pdf_obj.pages[r].extract_words()
            start = 0
            end = 0
            dict_packing = {}
            dict_values = {}
            # print(dict_words)
            # break
            text_words = self.pdf_obj.pages[r].extract_text()
            # print(text_words)
            Invo = ""


            for item in dict_words:
                if item["text"] == "INVOICE":
                    c = "INVOICE"
                    try:
                        pattarn = re.compile(r"INV. NO:(.*?)\n", re.S)
                        text_list = re.findall(pattarn, text_words)
                        Invo = text_list[-1].replace(" ", "")
                        pattarn = re.compile(r"(COO:|Made)(.*?)LINE", re.S)
                        made_list = re.findall(pattarn, text_words)
                        if made_list:
                            for item in made_list:
                                if item[1][3] == " ":
                                    Invo_coo_list.append(item[1][4:].replace(",", " ").strip())
                                else:
                                    Invo_coo_list.append(item[1].replace(",", " ").strip())
                        # print(Invo_coo_list)
                    except:
                        print("正则提取异常！")
                    # sys.exit()
                    self.get_invoice(dict_values, dict_words, end, config_info, c, Invo)
                    onepage_df = pd.DataFrame()
                    for title in dict_values:
                        onetitle_df = pd.DataFrame(dict_values.get(title), columns=[title])
                        onepage_df = pd.concat([onepage_df, onetitle_df], axis=1)
                    Invoice_Df = Invoice_Df.append(onepage_df)

                    for i in range(len(onepage_df)):
                        Invo_listone.append(Invo)

                    break
                elif item["text"] == "PACKING":
                    c = "PACKING"
                    try:
                        pattarn = re.compile(r"INV. NO:(.*?)\n", re.S)
                        text_list = re.findall(pattarn, text_words)
                        if "境外品牌" in text_words:
                            Invo = text_list[-1].replace(" ", "")
                            # Invo_list.append(Invo)
                        pattarn = re.compile(r"境外品牌(.*?)ECCN", re.S)
                        made_list = re.findall(pattarn, text_words)
                        for item in made_list:
                            if "Made" in item:
                                coo_list.append(item[13:].replace(" ", "").replace(",", " ").strip())
                            if "C.O.O" in item:
                                coo_list.append(item.split(":")[1].strip())
                    except:
                        print("正则提取异常！")
                    self.get_invoice(dict_packing, dict_words, end, config_info, c, Invo)
                    onepage_df = pd.DataFrame()
                    for title in dict_packing:
                        onetitle_df = pd.DataFrame(dict_packing.get(title), columns=[title])
                        onepage_df = pd.concat([onepage_df, onetitle_df], axis=1)
                    Packing_Df = Packing_Df.append(onepage_df)
                    for i in range(len(onepage_df)):
                        Invo_list.append(Invo)
                    break

        try:
            Invoice_Df["COO"] = Invo_coo_list
            Invoice_Df["Invo"] = Invo_listone
            Packing_Df["COO"] = coo_list
            Packing_Df["Invo"] = Invo_list
        except:
            print("COO未提取成功！")

        # print(dict_values)
        if Invoice_Df.empty:
            print("Invoice_Df为空！")
        else:
            Customer = []
            Brand = []
            Description = []
            for index, row in Invoice_Df.iterrows():
                Customer.append(row["Item"].rsplit(" ", 1)[1])
                row["Item"] = row["Item"].rsplit("(", 1)[0].strip()
                Brand.append(row["No."].split(" ", 1)[0])
                if "COO" in row["No."]:
                    Description.append(row["No."].split("COO")[0].strip().split(" ", 1)[1])
                if "Made" in row["No."]:
                    Description.append(row["No."].split("Made")[0].strip().split(" ", 1)[1])

            Invoice_Df["Customer"] = Customer
            Invoice_Df["Brand"] = Brand
            Invoice_Df["Description"] = Description
            Invoice_Df.to_excel(r"../ExcelFile/INVOICE.xlsx", index=False)
            ExcelComparison().invoke_pandas(r"../ExcelFile/INVOICE.xlsx")

        if Packing_Df.empty:
            print("Packing_Df为空！")
        else:
            Packing_Df.to_excel(r"../ExcelFile/PACKING.xlsx", index=False)
            ExcelComparison().invoke_pandas(r"../ExcelFile/PACKING.xlsx")

    def get_invoice(self, dict_values, dict_words, end, config_info, c, Invo):
        isExist = False
        for item in dict_words:
            if item["text"] == config_info[self.company_number]["Meau"][c]["end_first"]:
                end = dict_words.index(item)
                isExist = True
                break
        if not isExist:
            end = dict_words.index(dict_words[-1])
        list_columns = config_info[self.company_number]["Meau"][c]['columns']
        for r in list_columns:
            info_list = []
            for item in dict_words:
                if item["text"] == r:
                    start = dict_words.index(item)
                    # print(start)
                    x0 = dict_words[start]["x0"]
                    x1 = dict_words[start]["x1"]
                    top = dict_words[start]["top"]
                    bottom = dict_words[start]["bottom"]

                    x0_left = config_info[self.company_number]["Meau"][c]["columns"][r]["x0_left"]
                    x1_right = config_info[self.company_number]["Meau"][c]["columns"][r]["x1_right"]
                    top_0 = config_info[self.company_number]["Meau"][c]["columns"][r]["top_0"]
                    bottom_0 = config_info[self.company_number]["Meau"][c]["columns"][r]["bottom_0"]
                    desc_l = config_info[self.company_number]["Meau"][c]["columns"][r]["desc_l"]
                    desc_r = config_info[self.company_number]["Meau"][c]["columns"][r]["desc_r"]
                    move = config_info[self.company_number]["Meau"][c]["columns"][r]["move"]
                    move_second = config_info[self.company_number]["Meau"][c]["move_second"]

                    target = Controller().get_info(start, end, dict_words, x0, top, bottom, x1, info_list, x0_left,
                                                   x1_right, top_0,
                                                   bottom_0, move, r)
                    list_target = list(target)

                    if len(list_target) == 1:
                        Controller().get_other(dict_words[list_target[0]], dict_words, move_second,
                                               dict_values, r, end, start, info_list, top_0, bottom_0, desc_l,
                                               desc_r)
                    elif len(list_target) > 1:
                        list_ins = []
                        list_ins.append(" ".join(info_list))
                        info_list = list_ins
                        Controller().get_other(dict_words[list_target[0]], dict_words, move_second,
                                               dict_values, r, end, start, info_list, top_0, bottom_0, desc_l,
                                               desc_r)
                    elif len(list_target) == 0:
                        info = {"text": " ", "x0": dict_words[start]["x0"] - Decimal(x0_left),
                                "x1": dict_words[start]["x1"],
                                "top": dict_words[start]["top"] + Decimal(move),
                                "bottom": dict_words[start]["bottom"] + Decimal(move)}
                        info_list.append(" ")
                        Controller().get_other(info, dict_words, move_second,
                                               dict_values, r, end, start, info_list, top_0, bottom_0, desc_l,
                                               desc_r)

                    break

        return c


class Controller:

    def get_info(self, start, end, dict_words, x0, top, bottom, x1,
                 info_list, x0_left, x1_right, top_0, bottom_0, move, r):
        for i in range(start, end):
            if x0 - Decimal(x0_left) <= dict_words[i]["x0"] and dict_words[i]["x1"] <= x0 + Decimal(x1_right) and \
                    top + Decimal(move) - Decimal(top_0) <= dict_words[i]["top"] <= top + Decimal(
                move) + Decimal(top_0) and bottom + Decimal(move) - Decimal(bottom_0) <= dict_words[i][
                "bottom"] <= bottom + Decimal(move) + Decimal(bottom_0) and top + Decimal(move) < dict_words[end][
                "top"]:
                info_list.append(dict_words[i]["text"].replace("kg", ""))
                yield i

    def get_other(self, info, dict_words, move_second, dict_values, r, end, start, info_list, top_0,
                  bottom_0, desc_l, desc_r):
        x0 = info["x0"]
        top = info["top"]
        x1 = info["x1"]
        bottom = info["bottom"]
        for t in range(1, 100):
            top += Decimal(move_second)
            bottom += Decimal(move_second)
            if top > dict_words[end]["top"]:
                break
            isExist = False
            list_ins = []
            for c in range(start, end):
                if x0 - Decimal(desc_l) <= dict_words[c]["x0"] and dict_words[c]["x1"] <= x0 + Decimal(
                        desc_r) and top - Decimal(
                    top_0) <= dict_words[c]["top"] <= top + Decimal(top_0) and bottom - Decimal(bottom_0) <= \
                        dict_words[c]["bottom"] <= bottom + Decimal(bottom_0) and top + Decimal("10") < \
                        dict_words[end][
                            "bottom"]:
                    list_ins.append(dict_words[c]["text"].replace("kg", ""))
                    isExist = True
                    # break
            if not isExist:
                list_ins.append(" ")
            # print(list_ins)
            if len(list_ins) > 1:
                info_list.append(" ".join(list_ins))
            else:
                info_list.append(list_ins[0])

        # print(info_list)
        self.deal_none(info_list, dict_values, r)
        # print(dict_values)
        # sys.exit()

    def deal_none(self, info_list, dict_packing, r):
        cut_len = len(info_list)
        for null_len in range(len(info_list), -1, -1):
            if ' ' * null_len == ''.join(info_list[-null_len:]):
                cut_len = -null_len
                break
        content_list = info_list[:cut_len]
        for item in content_list:
            dict_packing.setdefault(r, []).append(item)
