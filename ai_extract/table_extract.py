import re
import sys


def each_keyword(re_list, txt_coord):
    """
    提取关键字的内容
    :param re_list: 配置中keywords需要提取的list
    :param txt_coord: 带位置的一页坐标
    :return: 所有匹配上的一个list list里面每一个元素是字典
    """
    # 存储关键字内容的list
    con_list = []
    for re_con in re_list:
        re_field = re.compile(re_con)
        for dict_str_coord in txt_coord:
            if re_field.search(dict_str_coord['text']):
                con_list.append(dict_str_coord)
    return con_list



def extract_conxy_coord(field, keywords_con_dict, txt_coord, bottom_keywords_con_list):
    """
    主要功能，提取内容
    :param field: 一个字段的配置 类型 dict
    :param keywords_con_dict: 提取出来关键字的list
    :param txt_coord: 带坐标的内容
    :param bottom_keywords_con_list: 底部关键字的内容
    :return: content_list 返回符合条件的内容 list里面是dict
    """
    content_list = []
    keyword_x = keywords_con_dict.get('x0') + field['move_x']
    keyword_y = keywords_con_dict.get('top') + field['move_y']
    isOnly = False
    if 'isOnly' in field.keys():
        isOnly = field['isOnly']
    # 提取第一个内容，后面判断他是不是多个值
    for con in txt_coord:
        # 内容提取的主要逻辑
        # 文本的x0 要比关键字的x0要大 and 文本的top要 比关键字的top大 比关键字的top加上高度y and 内容的x1要比关键x1加上x的值小
        if con.get('x0') > keyword_x and keyword_y < con.get('top') < keyword_y + field['y'] \
                and con.get('x1') < keyword_x + field['x']:
            content_list.append(con)
            if isOnly:
                break
    # 一个为空值给他伪造一个内容，针对第一个就为空值 多个值才会有这种情况
    if len(content_list) == 0 and 'next_con' in field:
        content_list.append({'text': ' ', 'top': keyword_y + field['y'] - 10})
    try:
        bottom_keywords_y = bottom_keywords_con_list[field['bottom_keywords_index']]
    except Exception as e:
        print(e, '没有bottom_keywords')

    # 容错
    try:
        next_top = content_list[-1].get('top')
        # print(next_top)
        # 对内容进行判断是否超出范围
        try:
            if next_top >= bottom_keywords_y.get('top'):
                content_list = []
                return content_list
        except Exception as e:
            print(e, 'bottom_keywords_y为1 只有一个值')
        # 如果存在 next_con  bottom_keywords bottom_keywords_index 就是多值
        if 'next_con' in field and len(bottom_keywords_con_list) != 0 and bottom_keywords_con_list[0] != 1:
            # 添加多个内容
            for i in range(100):
                next_top += field['next_con']
                # print(keyword_x)
                # 超过bottom_keywords_y 底部关键字的top值后退出
                if next_top > bottom_keywords_y.get('top'):
                    break
                next_tag = True
                for con in txt_coord:
                    # print('jin')
                    # 文本的x0 要比关键字的x0要大 and
                    # 文本的top要 比关键字的top大 比关键字的top加上高度y and
                    # 内容的x1要比关键x1加上x的值小 and
                    # 内容的top要比底部关键字top加上x的值小

                    if con.get('x0') > keyword_x and next_top <= con.get('top') <= next_top + field['y'] \
                            and con.get('x1') < keyword_x + field['x'] and \
                            next_top + field['y'] < bottom_keywords_y.get('top'):
                        next_tag = False
                        content_list.append(con)
                if next_tag:
                    content_list.append({'text': ' ', 'top': next_top})
        # 只有一个值
        else:
            return content_list
    except Exception as e:
        print(e, 'content_list[-1]找不到内容')
    # 将不必要空值进去删除
    cut_len = len(content_list)
    txt_list = [i.get('text') for i in content_list]
    for null_len in range(len(txt_list), -1, -1):
        if ' ' * null_len == ''.join(txt_list[-null_len:]):
            cut_len = -null_len
            break
    content_list = content_list[:cut_len]
    return content_list


def keywords_extract(field, txt_coord):
    """

    :param field: 一个配置的信息
    :param txt_coord: 带坐标的内容
    :return: content_list list里面是dict
    """
    # 加上index就可以拿到哪个re匹配出来对的坐标了,field为每个字段
    keywords_con_list = []
    bottom_keywords_con_list = []
    # 对关键字进行提取
    for re_con in field['keywords']:
        re_field = re.compile(re_con)
        # 遍历 txt_coord
        for dict_str_coord in txt_coord:

            if re_field.search(dict_str_coord['text']):
                keywords_con_list.append(dict_str_coord)
    # 拿到了关键字的坐标后，拿坐标的关键字
    # if 'x' in field and 'y' in field:
    try:

        if 'bottom_keywords' in field:
            # 这里就是多值提取
            bottom_keywords_con_list.extend(each_keyword(field['bottom_keywords'], txt_coord))
            content_list = extract_conxy_coord(field, keywords_con_list[field['index']], txt_coord,
                                               bottom_keywords_con_list)
            # 将同一行的内容合并
            content_list = norm_output(content_list)
            return content_list
        else:
            # 只有一个值提取
            content_list = extract_conxy_coord(field, keywords_con_list[field['index']], txt_coord, [1])
            # 将同一行的内容合并
            content_list = norm_output(content_list)
            return content_list
    except Exception as e:
        print(e, "该文本没有这个字段")


def norm_output(content_list):
    """
    :param content_list:[{},{}]
    :return:{field: {}}
    """
    lines = {}
    input_list = []
    for each_content in content_list:
        filter_data = filter(lambda x: x['top'] == each_content['top'] and x['x0'] == each_content['x0'], input_list)
        if len(list(filter_data)) > 0:
            continue
        if str(each_content.get('top')) in lines:
            # top相同就合并成一行
            lines[str(each_content.get('top'))] += " " + each_content.get('text')
            input_list.append(each_content)
        else:
            lines[str(each_content.get('top'))] = each_content.get('text')
            input_list.append(each_content)
    # sys.exit()
    return lines


def area_extract(table_re_setting, txt_coord):
    """

    :param table_re_setting: 表格提取的配置
    :param txt_coord: PDF中所有临近字段的坐标
    :return:
    """
    # 存储配置需要提取的所有字段
    all_field = {}
    for field in table_re_setting:
        # field是要抽取的字段名称格式ci：{每个字段：配置}
        # 将每一个配置信息传进keywords_extract
        content_list = keywords_extract(table_re_setting[field], txt_coord)
        # sys.exit()
        try:
            # 将提取出来的内容进行内容提取
            con_list = [i for i in content_list.values()]
            all_field.update({field: con_list})
        except Exception as e:
            print(field)
            print(e, "area_extract error: [i.get('text') for i in content_list]")
    return all_field
