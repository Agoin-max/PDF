import pandas as pd
import pdfplumber
# from dispose_method import data_processing
# from table_extract import area_extract
# from table_field_setting import jre_table
from Supplier_ARui.Comparison_In_Pk import ExcelComparison
from ai_extract import dispose_method, table_extract, table_field_setting





def extract_all(key, path):
    """
    提取入口
    :param path: 文件路径
    :return: 文件中需要抽取的所有字段
    """
    pdf = pdfplumber.open(path)
    INVOICE_df = pd.DataFrame()
    PACKING_LIST_df = pd.DataFrame()
    # 拿到每一页内容
    for index_page, page in enumerate(pdf.pages):
        inv_index = page.extract_text().upper().replace(' ', '').find('INVOICE')
        pak_index = page.extract_text().upper().replace(' ', '').find('PACKINGLIST')
        # print(page.extract_text())
        if (-1 < inv_index < pak_index) or (pak_index == -1 and inv_index > -1):
            x_tolerance = 3
            if 'inv_x_tolerance' in table_field_setting.jre_table[key].keys():
                x_tolerance = table_field_setting.jre_table[key]['inv_x_tolerance']
            # 带位置的内容
            txt_coord = page.extract_words(x_tolerance=x_tolerance, y_tolerance=3, keep_blank_chars=False,
                                           use_text_flow=False, horizontal_ltr=True, vertical_ttb=True,
                                           extra_attrs=[])
            for item in txt_coord:
                print(item)
            # 提取INVOICE
            res_iv = table_extract.area_extract(table_field_setting.jre_table[key]['iv'], txt_coord)
            onepage_df = pd.DataFrame()
            # 将返回的dict转为datafrmae进行合并 pd.concat
            for title in res_iv:
                onetitle_dataframe = pd.DataFrame(res_iv.get(title), columns=[title])
                onepage_df = pd.concat([onepage_df, onetitle_dataframe], axis=1)
            # 将一页的内容添加到INVOICE_df
            INVOICE_df = INVOICE_df.append(onepage_df)
        # PACKING LIST 提取内容
        elif (-1 < pak_index < inv_index) or (inv_index == -1 and pak_index > -1):
            x_tolerance = 3
            if 'pak_x_tolerance' in table_field_setting.jre_table[key].keys():
                x_tolerance = table_field_setting.jre_table[key]['pak_x_tolerance']
            # 带位置的内容
            txt_coord = page.extract_words(x_tolerance=x_tolerance, y_tolerance=3, keep_blank_chars=False,
                                           use_text_flow=False, horizontal_ltr=True, vertical_ttb=True,
                                           extra_attrs=[])
            # 提取PACKING LIST
            res_iv = table_extract.area_extract(table_field_setting.jre_table[key]['pl'], txt_coord)
            onepage_df = pd.DataFrame()
            # 将返回的dict转为datafrmae进行合并 pd.concat
            for title in res_iv:
                onetitle_dataframe = pd.DataFrame(res_iv.get(title), columns=[title])
                onepage_df = pd.concat([onepage_df, onetitle_dataframe], axis=1)
            # 将一页的内容添加到PACKING_LIST_df
            if len(PACKING_LIST_df) > 0:
                PACKING_LIST_df = PACKING_LIST_df.append(onepage_df)
            else:
                PACKING_LIST_df = onepage_df

    pdf.close()

    return [PACKING_LIST_df, INVOICE_df]


if __name__ == '__main__':
    # company_name 哪一个公司 file_path就是文件路径
    company_name = "airui"
    file_name = "PASDR_0107857_45.pdf"
    file_path = r'./files/{}/{}'.format(company_name, file_name)
    PACKING_LIST_df, INVOICE_df = extract_all(company_name, file_path)
    PACKING_LIST_df, INVOICE_df = dispose_method.data_processing(company_name, PACKING_LIST_df, INVOICE_df)

    print("-" * 100)
    print(PACKING_LIST_df)
    print("-" * 100)
    print(INVOICE_df)

    # 生成excel
    PACKING_LIST_df.to_excel('./ExcelList/PACKING.xlsx', index=0)

    INVOICE_df.to_excel('./ExcelList/INVOICE.xlsx', index=0)


def invoke_main(company_name, file_path):
    # company_name 哪一个公司 file_path就是文件路径
    # file_path = r'./files/{}/{}'.format(company_name, file_name)
    PACKING_LIST_df, INVOICE_df = extract_all(company_name, file_path)
    PACKING_LIST_df, INVOICE_df = dispose_method.data_processing(company_name, PACKING_LIST_df, INVOICE_df)
    # 生成excel

    if PACKING_LIST_df.empty:
        print("PACKING_LIST_df为空！")
    else:
        PACKING_LIST_df.to_excel('../ExcelFile/PACKING.xlsx', index=0)
        ExcelComparison().invoke_pandas('../ExcelFile/PACKING.xlsx')

    if INVOICE_df.empty:
        print("INVOICE_df为空！")
    else:
        INVOICE_df.to_excel('../ExcelFile/INVOICE.xlsx', index=0)
        ExcelComparison().invoke_pandas('../ExcelFile/INVOICE.xlsx')
