import os
import sys


class FuTools:
    def open_json(self):
        sys.path.append(os.getcwd())
        f = open("../Setting/setting.json", "r", encoding="utf-8")
        return f.read()
