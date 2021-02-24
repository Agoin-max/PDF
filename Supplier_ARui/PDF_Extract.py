from decimal import Decimal
import pandas as pd
import json
from Setting.Father import DataPdf
from Setting.Tools import FuTools
import sys
import re

from Supplier_HongYi.Comparison_In_Pk import ExcelComparison
from ai_extract.all_extract import invoke_main


class ExtractAirRuiPdf(DataPdf):

    def __init__(self, pdf_pages=0, pdf_obj=object, company_number="", path=""):
        self.pdf_pages = pdf_pages
        self.pdf_obj = pdf_obj
        self.company_number = company_number
        self.path = path

    def get_info(self):
        # print(self.path)
        # sys.exit()
        invoke_main(self.company_number, self.path)

    def get_invoice(self, dict_values, dict_words, end, config_info, c):
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
                    desc = config_info[self.company_number]["Meau"][c]["columns"][r]["desc"]
                    move = config_info[self.company_number]["Meau"][c]["columns"][r]["move"]
                    move_second = config_info[self.company_number]["Meau"][c]["move_second"]

                    target = Controller().get_info(start, end, dict_words, x0, top, bottom, x1, info_list, x0_left,
                                                   x1_right, top_0,
                                                   bottom_0, move)
                    list_target = list(target)

                    if len(list_target) == 1:
                        Controller().get_packing_other(dict_words[list_target[0]], dict_words, move_second,
                                                       dict_values, r, end, start, info_list, top_0, bottom_0, desc)
                    elif len(list_target) > 1:
                        list_ins = []
                        list_ins.append("".join(info_list))
                        info_list = list_ins
                        Controller().get_packing_other(dict_words[list_target[0]], dict_words, move_second,
                                                       dict_values, r, end, start, info_list, top_0, bottom_0, desc)
                    elif len(list_target) == 0:
                        info = {"text": " ", "x0": dict_words[start]["x0"] - Decimal(x0_left),
                                "x1": dict_words[start]["x1"],
                                "top": dict_words[start]["top"] + Decimal(move),
                                "bottom": dict_words[start]["bottom"] + Decimal(move)}
                        info_list.append(" ")
                        Controller().get_packing_other(info, dict_words, move_second,
                                                       dict_values, r, end, start, info_list, top_0, bottom_0, desc)

                    break

        return c


class Controller:

    def get_info(self, start, end, dict_words, x0, top, bottom, x1,
                 info_list, x0_left, x1_right, top_0, bottom_0, move):
        for i in range(start, end):
            if x0 - Decimal(x0_left) <= dict_words[i]["x0"] and dict_words[i]["x1"] <= x1 + Decimal(x1_right) and \
                    top + Decimal(move) - Decimal(top_0) <= dict_words[i]["top"] <= top + Decimal(
                move) + Decimal(top_0) and bottom + Decimal(move) - Decimal(bottom_0) <= dict_words[i][
                "bottom"] <= bottom + Decimal(move) + Decimal(bottom_0) and top + Decimal(move) < dict_words[end][
                "top"]:
                info_list.append(dict_words[i]["text"].replace("kg", ""))
                yield i

    def get_packing_other(self, info, dict_words, move_second, dict_values, r, end, start, info_list, top_0,
                          bottom_0, desc):
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
                if x0 - Decimal(desc) <= dict_words[c]["x0"] and dict_words[c]["x1"] <= x1 + Decimal(
                        desc) and top - Decimal(
                    top_0) <= dict_words[c]["top"] <= top + Decimal(top_0) and bottom - Decimal(bottom_0) <= \
                        dict_words[c]["bottom"] <= bottom + Decimal(bottom_0) and top + Decimal("10") < \
                        dict_words[end][
                            "bottom"]:
                    list_ins.append(dict_words[c]["text"].replace("kg", ""))
                    isExist = True
                    break
            if not isExist:
                list_ins.append(" ")
            # print(list_ins)
            if len(list_ins) > 1:
                info_list.append("".join(list_ins))
            else:
                info_list.append(list_ins[0])

        # print(info_list)
        self.deal_none(info_list, dict_values, r)
        # print(dict_values)
        # sys.exit()

    def deal_none(self, info_list, dict_values, r):
        cut_len = len(info_list)
        for null_len in range(len(info_list), -1, -1):
            if ' ' * null_len == ''.join(info_list[-null_len:]):
                cut_len = -null_len
                break
        content_list = info_list[:cut_len]
        for item in content_list:
            dict_values.setdefault(r.replace(".", "").replace("'", "").upper(), []).append(item)
