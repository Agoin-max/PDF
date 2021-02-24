import os
import sys
import pandas as pd
import pdfplumber
import json
from Setting.Tools import FuTools
from Supplier_HongYi.Doc_Pdf import createPdf
from Supplier_ShiPing.PDF_Extract import ExtractPdfWorld
from Supplier_WenYe.PDF_Extract import ExtractPdf
from Supplier_XinTai.PDF_Extract import ExtractXinTaiPdf
from Supplier_FengYi.PDF_Extract import ExtractFengYiPdf
from Supplier_HongYi.PDF_Extract import ExtractHongYiPdf
from Supplier_ARui.PDF_Extract import ExtractAirRuiPdf
from Supplier_ZengNiQiang.PDF_Extract import ExtractZengNiQiangPdf
from Supplier_Samtec.PDF_Extract import ExtractSamtecPdf
from Supplier_JunLong.PDF_Extract import ExtractJunLongPdf
from Supplier_PingJia.PDF_Extract import ExtractPingJiaWorld


class BasePdf:

    def __init__(self, path=""):
        self.path = path

    def open_pdf(self):
        """path:PDF路径"""
        return pdfplumber.open(self.path)

    def get_page(self):
        """获取PDF页数"""
        return len(self.open_pdf().pages)


class DealPdf:
    def __init__(self, path, company_number=""):
        if path.find("doc") >= 0:
            path = createPdf(path)
        self.pdf_obj = BasePdf(path).open_pdf()
        self.pdf_pages = BasePdf(path).get_page()
        self.company_number = company_number
        self.path = path
        self.__config_file = json.loads(FuTools().open_json())

    def info_extract(self):
        eval(self.__config_file[self.company_number]["Name"]).get_info()
        self.pdf_obj.close()


class Enter:
    def get_file_path(self, number):
        list_file = os.listdir("D:\\PDF\\" + number)
        for item in list_file:
            yield "D:\\PDF\\" + number + "\\" + item


DealPdf(r"C:\Users\windo\Desktop\PL_64385465_G321020405.pdf", "0114").info_extract()
