import pymysql
import re
import time
import pandas as pd


class MysqlOperation:
    def __init__(self):
        self.kwargs = {
            "host": "192.168.0.178",
            "port": 3306,
            "user": "root",
            "password": "root",
            "database": "gw",
            "charset": "utf8"
        }
        self.connect()

    def connect(self):
        self.db = pymysql.connect(**self.kwargs)
        self.cur = self.db.cursor()

    def close(self):
        self.cur.close()
        self.db.close()

    def insert_info(self, iterable, sql):
        try:
            self.cur.executemany(sql, iterable)
            self.db.commit()  # 提交事务
        except Exception as e:
            print(e)
            self.db.rollback()  # 事务回滚

    def find_info(self, sql):
        self.cur.execute(sql)  # 执行sql语句
        info = self.cur.fetchall()
        if info:
            return list(info)
        else:
            return "未查询到结果！"

    def update_info(self, sql):
        try:
            self.cur.execute(sql)
            self.db.commit()
            print("修改成功！")
        except Exception as e:
            print(e)
            self.db.rollback()

    def delete_info(self):
        pass


class DealController:

    def deal_info(self, ins, **kwargs):
        keyNo = time.strftime("%Y%m%d%H%M%S")
        InserTime = time.strftime("%Y-%m-%d %H-%M-%S")
        kwargs["keyNo"] = keyNo
        kwargs["is_lock"] = 0
        kwargs["state"] = 0
        kwargs["InserTime"] = InserTime
        words = [tuple(kwargs.values())]
        sql = "insert into cassify (CompanyName, ContractNo, BusinessName, guid, FileName, keyNo, is_lock, state, InserTime) values (%s,%s,%s,%s,%s,%s,%s,%s,%s);"
        ins.insert_info(words, sql)
        df = pd.read_excel("D:\\Mail\\sendexcel\\" + kwargs["FileName"])
        df = df.where(df.notnull(), "")
        new_list = [tuple([keyNo, int(row["序号"]), str(row["货物中文名称"]), str(row["品牌"]), str(row["型号"]), str(row["单位"]), str(row["商品描述"]), str(row["报关品名"]), str(row["用途"]), str(row["变动要素1"]), str(row["变动要素2"]), str(row["变动要素3"]), str(row["税号"]), str(row["规格型号"])]) for index, row in df.iterrows()]
        sql = "insert into cassifydetail (keyNo,S_NO, ChineseName, Brand, Model, Unit, Description, DeclaratioNname, UseName,Change1,Change2,Change3,TaxNumber,SpecificationModel) values (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s);"
        ins.insert_info(new_list, sql)


if __name__ == '__main__':
    ins = MysqlOperation()
    # DealController().deal_info(ins, CompanyName="深圳市思维颗粒科技有限公司", ContractNo="没有查到相应数据", BusinessName="商务部凌宝艳", guid="ccfdafd5-e640-480e-89a9-39be8303aeab", FileName="018129553_20210205_110702_newtable.xlsx")
    print(ins.find_info(sql="select * from cassify where state = 1"))
    # ins.update_info(sql="update cassify set state = 2 where keyNo = '20210208164359'")
    ins.close()
