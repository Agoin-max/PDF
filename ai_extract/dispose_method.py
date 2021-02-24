import re

import pandas as pd


def junlong_method(PACKING_LIST_df, INVOICE_df):
    """骏龙数据处理"""
    # if len(str(PACKING_LIST_df.loc[0]["Part"])) > 1:
    if len(PACKING_LIST_df) > 0:
        # 将所有的nan替换为空字符串
        PACKING_LIST_df.fillna("", inplace=True)
        # 循环去重
        for index, row in PACKING_LIST_df.iterrows():
            comments_str = str(row["Comments"])
            part_str = str(row["Part"])
            if len(comments_str) > 1:
                # 以空行切割成列表
                temp_list = comments_str.split(" ")
                row["Comments"] = temp_list[0]

            if len(part_str) > 1:
                # 以空行切割成列表
                temp_list = part_str.split(" ")
                row["Part"] = temp_list[0]

        # 取出Part列需要的数据
        part_list = list(PACKING_LIST_df["Part"])
        # 取出Comment列需要的值
        comments_list = list(PACKING_LIST_df["Comments"])
        new_comments_list = list()
        if len(comments_list[-1]) > 1:
            comments_list.append("")
        if len(part_list[-1]) > 1:
            part_list.append("")

        for i, j in enumerate(comments_list):
            if len(j) <= 1:
                print(j)
                if len(comments_list[i - 1]) > 1:
                    new_comments_list.append(comments_list[i - 1])
                part_list[i] = ""

        new_part_list = list()
        flag_index = 0
        for i, j in enumerate(part_list):
            if len(j) <= 1:
                if len(part_list[i - 1]) > 1:
                    while True:
                        if len(part_list[flag_index]) <= 1:
                            flag_index += 1
                        else:
                            break
                    dispose_part_list = [part_list[x] for x in range(flag_index, i)]
                    # print(dispose_part_list)
                    dispose_part_list = dispose_part_list[:2]
                    if len(dispose_part_list) >= 2:
                        if dispose_part_list[0] != dispose_part_list[1]:
                            for m, k in enumerate(dispose_part_list):
                                if k.endswith("-") and m < 1:
                                    new_part_list.append(k + dispose_part_list[m + 1])
                                    break
                                elif k.startswith("-"):
                                    new_part_list.append(dispose_part_list[m - 1] + k)
                                    break
                            else:
                                new_part_list.append(dispose_part_list[0])
                        else:
                            new_part_list.append(dispose_part_list[0])
                    else:
                        new_part_list.append(dispose_part_list[0])
                    flag_index = i
        flag_index = 0
        # 保存筛选后的数据
        data_dict = {
            "Packing_List": [], "CT": [], "Comments": [], "Part": [],
            "Total_Quantity": [], "Country_of_Origin": [], "Net": [], "Gross": []
        }
        # 循环添加数据
        for index, row in PACKING_LIST_df.iterrows():
            if len(row["Total_Quantity"]) > 1:
                pk = PACKING_LIST_df.loc[0]["Packing_List"]
                pk = pk if isinstance(pk, str) else list(pk)[0]
                data_dict["Packing_List"].append(pk)
                data_dict["CT"].append(row["CT"])
                data_dict["Comments"].append(new_comments_list[index])
                data_dict["Part"].append(new_part_list[index])
                data_dict["Total_Quantity"].append(str(row["Total_Quantity"]).replace(',', ''))
                data_dict["Country_of_Origin"].append(row["Country_of_Origin"])
                data_dict["Net"].append(row["Net"])
                data_dict["Gross"].append(row["Gross"])
        # print(data_dict)
        PACKING_LIST_df = pd.DataFrame(data_dict)
        PACKING_LIST_df = PACKING_LIST_df.rename(
            columns={'Packing_List': 'Invo', 'Total_Quantity': 'QTY', 'Part': 'PN'})

    if len(INVOICE_df) > 0:
        # 保存筛选后的数据
        data_dict = {
            "Invoice_Number": [], "Customer_PO": [], "Customer_P/N": [],
            "End_Customer_P/N": [], "Quantity": [], "Unit_Price": []
        }
        # 将所有的nan替换为空字符串
        INVOICE_df.fillna("", inplace=True)
        # 循环去除空数据
        for index, row in INVOICE_df.iterrows():
            if len(row["Customer_PO"]) <= 1 and len(row["Description"]) <= 1:
                INVOICE_df.drop(index, axis=0, inplace=True)
        # 循环添加数据
        for index, row in INVOICE_df.iterrows():
            if len(row["Quantity"]) > 1:
                data_dict["Invoice_Number"].append(INVOICE_df.loc[0]["Invoice_Number"])
                data_dict["Customer_PO"].append(INVOICE_df.loc[index + 1]["Customer_PO"])
                data_dict["Customer_P/N"].append(row["Description"])
                data_dict["End_Customer_P/N"].append(INVOICE_df.loc[index + 1]["Description"].replace("-", ""))
                data_dict["Quantity"].append(str(row["Quantity"]).replace('PCS', '').replace(',', '').strip())
                data_dict["Unit_Price"].append(row["Unit_Price"])
        print(data_dict)
        INVOICE_df = pd.DataFrame(data_dict)
        INVOICE_df = INVOICE_df.rename(columns={'Invoice_Number': 'Invo', 'Quantity': 'QTY', 'Customer_P/N': 'PN'})
    return PACKING_LIST_df, INVOICE_df


def airui_method(PACKING_LIST_df, INVOICE_df):
    """艾睿数据处理"""
    data_dict = {
        "CUSTOMER_ORDER_NUMBER": [], "Invo": [], "CARTON": [],
        "DIMENSION": [], "GW": [], "NW": [], "PN": [], "CPN": [], "COO": [], "QTY": []
    }

    # 将所有的nan替换为空字符串
    PACKING_LIST_df.fillna("", inplace=True)

    # 定义变量保存填充值
    carton = ""
    dimension = ""
    # 循环添加有效数据
    for index, row in PACKING_LIST_df.iterrows():
        if row["COO"] and row["CPN"] and row["QTY"] and row["PN"]:
            cn = PACKING_LIST_df.loc[0]["CUSTOMER_ORDER_NUMBER"]
            cn = cn if isinstance(cn, str) else list(cn)[0]
            data_dict["CUSTOMER_ORDER_NUMBER"].append(cn)
            ivn = PACKING_LIST_df.loc[0]["Invo"]
            ivn = cn if isinstance(ivn, str) else list(ivn)[0]
            data_dict["Invo"].append(ivn)
            data_dict["COO"].append(row["COO"])
            if row["CARTON"]:
                carton = row["CARTON"]
                data_dict["CARTON"].append(carton)
            else:
                data_dict["CARTON"].append(carton)
            if row["DIMENSION"]:
                dimension = row["DIMENSION"]
                data_dict["DIMENSION"].append(dimension)
            else:
                data_dict["DIMENSION"].append(dimension)
            data_dict["GW"].append(row["GW"])
            data_dict["NW"].append(row["NW"])
            data_dict["PN"].append(row["PN"])
            data_dict["CPN"].append(row["CPN"])
            data_dict["QTY"].append(row["QTY"])

    PACKING_LIST_df = pd.DataFrame(data_dict)

    return PACKING_LIST_df, INVOICE_df


def zenitron_inv(INVOICE_df):
    df = INVOICE_df
    # df = df.loc[(pd.notna(df['PO']) & pd.notna(df['ITEM']))]
    # df = df.loc[~(pd.isna(df['ITEM']) & pd.isna(df['PO']))].reset_index(drop=True)
    df = df.loc[(df['ITEM'] != '') & (df['PO'] != ' ')].reset_index(drop=True)
    # df.to_excel(r"C:\Users\windo\Desktop\setting.xlsx")
    df['ITEM'].fillna('', inplace=True)
    df['PO'].fillna('', inplace=True)
    df['BRAND'].fillna('', inplace=True)
    df['ORIGIN'].fillna('', inplace=True)
    df['QTY'].fillna(method='pad', inplace=True)
    df['PRICE'].fillna(method='pad', inplace=True)
    df['Invo'].fillna(method='pad', inplace=True)
    res_df = pd.DataFrame([], columns=['PO', 'BRAND', 'DESCRIPTION', 'ITEM', 'ORIGIN', 'PART_NO', 'QTY', 'PRICE',
                                       'Invo'])
    item_list = []
    item = ''
    for index, row in df.iterrows():
        res = {}
        if str(row['PO']) != '' and index % 2 == 0:
            res['PO'] = str(row['PO']) + str(df.loc[index + 1, 'PO'])
            res['PART_NO'] = str(row['PART_NO']) + str(df.loc[index + 1, 'PART_NO'])
            rex = re.compile(r'^[A-Z]')
            if rex.search(str(df.loc[index + 1, 'ORIGIN'])) is not None:
                res['ORIGIN'] = str(row['ORIGIN']) + '/' + str(df.loc[index + 1, 'ORIGIN'])
            else:
                res['ORIGIN'] = str(row['ORIGIN'])
            res['QTY'] = row['QTY']
            res['PRICE'] = row['PRICE']
            res['Invo'] = row['Invo']
            brand = str(row['BRAND']) + str(df.loc[index + 1, 'BRAND'])
            datas = brand.split('/')
            res['BRAND'] = datas[0].strip()
            res['DESCRIPTION'] = datas[1].strip()
            if len(item.strip()) > 0:
                item_list.append(item)
                item = ''
            item = str(row['ITEM'])
            res_df = res_df.append(res, ignore_index=True)
        else:
            item += str(row['ITEM'])
        if (index + 1) == len(df):
            item_list.append(item)

    for index, row in res_df.iterrows():
        res_df.loc[index, 'ITEM'] = item_list[index]
    res_df = res_df.rename(columns={'ITEM': 'PN'})
    return res_df


def zenitron_pak(PACKING_LIST_df):
    df = PACKING_LIST_df
    # df = df.loc[(pd.notna(df['PO']) & pd.notna(df['ITEM']))]
    df['ITEM'].fillna(' ', inplace=True)
    df['QTY'].fillna(' ', inplace=True)
    df = df.loc[df['ITEM'] != ' '].reset_index(drop=True)
    df['INVOICE_NUMBER'].fillna(method='pad', inplace=True)
    res_df = pd.DataFrame([], columns=['CN_NO', 'ITEM', 'ORIGIN', 'QTY', 'NW', 'GW', 'INVOICE_NUMBER'])
    item_list = []
    item = ''
    for index, row in df.iterrows():
        res = {}
        if str(row['QTY']) != ' ':
            res['CN_NO'] = row['CN_NO']
            res['ORIGIN'] = row['ORIGIN']
            res['QTY'] = str(row['QTY']).replace("PCS", "").strip()
            res['NW'] = row['NW']
            res['GW'] = row['GW']
            res['INVOICE_NUMBER'] = row['INVOICE_NUMBER']
            if len(item.strip()) > 0:
                item_list.append(item)
                item = ''
            item = str(row['ITEM'])
            res_df = res_df.append(res, ignore_index=True)
        else:
            item += str(row['ITEM'])
        if (index + 1) == len(df):
            item_list.append(item)
    for index, row in res_df.iterrows():
        res_df.loc[index, 'ITEM'] = item_list[index].split('/')[0].strip()
    res_df = res_df.rename(columns={'ITEM': 'PN', 'INVOICE_NUMBER': 'Invo'})
    res_df = res_df.loc[res_df['PN'] != ''].reset_index(drop=True)
    return res_df

def fengyi_inv(INVOICE_df):
    # print(INVOICE_df)
    INVOICE_df['PO'].fillna(' ', inplace=True)
    df = INVOICE_df.loc[INVOICE_df['PO'] != ' '].reset_index(drop=True)
    df['Invo'].fillna(method='pad', inplace=True)
    res_df = pd.DataFrame([], columns=['PO', 'PN', 'Brand', 'QTY', 'Desc3', 'Cust', 'PRICE', 'Invo'])
    description_list = []
    for index, row in df.iterrows():
        if index % 2 == 0:
            data = {}
            data['PO'] = str(row['PO']) + str(df.loc[index + 1, 'PO'])
            description_list.append(str(row['DESCRIPTION']) + ' ' + str(df.loc[index + 1, 'DESCRIPTION']) + '|')
            data['QTY'] = row['QTY']
            data['PRICE'] = row['PRICE']
            data['Invo'] = row['Invo']
            res_df = res_df.append(data, ignore_index=True)

    pn_pattern = r'([\d\D]*?)\('
    brand_pattern = r'\(([\d\D]*?)IC'
    # print(description_list)
    for index, row in res_df.iterrows():
        description = description_list[index]
        pn_rex = re.search(pn_pattern, description)
        if pn_rex:
            res_df.loc[index, 'PN'] = pn_rex.group(1).strip()
        brand_rex = re.search(brand_pattern, description)
        if brand_rex:
            brand = brand_rex.group(1).replace('BRAND', '').strip()
            res_df.loc[index, 'Brand'] = brand
        desc_list = description.split(' ')
        for desc in desc_list:
            if ')' in desc:
                data = desc.split(')')
                res_df.loc[index, 'Desc3'] = data[0]
        pattern = re.compile(r"Cust.Part:(.*?)\|", re.S)
        # print(description)
        rex = re.findall(pattern, description)
        # print(rex)
        if len(rex) > 0:
            res_df.loc[index, 'Cust'] = rex[-1]
    return res_df


def data_processing(company_name, PACKING_LIST_df, INVOICE_df):
    """判断调用那个公司的数据处理方法"""
    if company_name == "0118":
        PACKING_LIST_df, INVOICE_df = airui_method(PACKING_LIST_df, INVOICE_df)
    elif company_name == "0121":
        PACKING_LIST_df, INVOICE_df = junlong_method(PACKING_LIST_df, INVOICE_df)
    elif company_name == '0119':
        if PACKING_LIST_df.empty:
            print("PACKING_LIST_df为空")
        else:
            PACKING_LIST_df = zenitron_pak(PACKING_LIST_df)
        if INVOICE_df.empty:
            print("INVOICE_df为空")
        else:
            INVOICE_df = zenitron_inv(INVOICE_df)
    elif company_name == '0112':
        if INVOICE_df.empty:
            print("INVOICE_df为空")
        else:
            INVOICE_df = fengyi_inv(INVOICE_df)


    return PACKING_LIST_df, INVOICE_df
