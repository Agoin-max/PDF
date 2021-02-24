from Setting.Tools import FuTools
import json
import os
import sys
from Supplier_WenYe.Comparison import DataComparison
from Supplier_ShiPing.Comparison import ShiPingComparison
from Supplier_XinTai.Comparison import XinTaiComparison
from Supplier_FengYi.Comparison import FengYiComparison
from Supplier_HongYi.Comparison import HongYiComparison
from Supplier_ARui.Comparison import AiRuiComparison
from Supplier_ZengNiQiang.Comparison import ZengNiQiangComparison
from Supplier_Samtec.Comparison import SamtecComparison
from Supplier_JunLong.Comparison import JunLongComparison
from Supplier_PingJia.Comparison import PingJiaComparison


class AllComparison:
    def __init__(self, company_number=""):
        self.company_number = company_number
        self.__config_file = json.loads(FuTools().open_json())
        sys.path.append(os.getcwd())

    def comp_all(self):
        return eval(self.__config_file[self.company_number]["Comparison"]).formation_excel()


print(AllComparison("0114").comp_all())
