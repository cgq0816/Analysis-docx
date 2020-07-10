# coding:utf-8
import requests
import docx
import json

temp_report_path = './docx_data/000858_4.txt'
nlp_url = 'http://hanlp-nlp-service:31001/swagger-ui.html#!/hanlp-controller/nlpSegmentUsingGET'

# text_dict = {"textbox_shape": {"width": "150.050pt", "margin-top": "179.299973pt", "margin-left": "42.599998pt", "height": "593.4pt", "mso-position-horizontal-relative": "page", "z-index": "-59416", "mso-position-vertical-relative": "page", "position": "absolute"}, "textbox_content": [{"font_size": "21", "text": "基础数据", "font": "楷体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "上证综指", "font": "楷体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "3086", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "总股本（万股）", "font": "楷体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "388161", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "已上市流通股（万股）", "font": "楷体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "379577", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "总市值（亿元）", "font": "楷体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "3970", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "流通市值（亿元）", "font": "楷体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "3882", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "每股净资产（", "font": "楷体"}, {"font_size": "18", "text": "MRQ", "font": "宋体"}, {"font_size": "18", "text": "）", "font": "楷体"}, {"font_size": "18", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "18.0", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "ROE", "font": "宋体"}, {"font_size": "18", "text": "（", "font": "楷体"}, {"font_size": "18", "text": "TTM", "font": "宋体"}, {"font_size": "18", "text": "）", "font": "楷体"}, {"font_size": "18", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "21.3", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "资产负债率", "font": "楷体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "%", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "主要股东", "font": "楷体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "宜宾市国有资产经营", "font": "楷体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "主要股东持股比例", "font": "楷体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "35.21%", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "21", "text": "股价表现", "font": "楷体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "%", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "21", "text": "1m", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "21", "text": "6m", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "21", "text": "12m", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "绝对表现", "font": "楷体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "23", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "21", "text": "87", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "21", "text": "44", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "相对表现", "font": "楷体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "18", "text": "18", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "21", "text": "64", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "21", "text": "40", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "16", "text": "(%)", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "14", "text": "五粮液", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "14", "text": "沪深300", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "15", "text": "60", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "15", "text": "40", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "15", "text": "20", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "15", "text": "0", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "15", "text": "-20", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "15", "text": "-40", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "15", "text": "Mar/18", "font": "楷体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "21", "text": "Jul/18", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "21", "text": "Nov/18", "font": "宋体"}, {"font_size": "21", "text": "tab/", "font": "宋体"}, {"font_size": "15", "text": "Mar/19 ", "font": "宋体"}, {"font_size": "18", "text": "资料来源：贝格数据、招商证券 ", "font": "楷体"}, {"font_size": "21", "text": "相关报告", "font": "楷体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "18", "text": "1", "font": "宋体"}, {"font_size": "18", "text": "、《五粮液", "font": "楷体"}, {"font_size": "18", "text": "（", "font": "楷体"}, {"font_size": "18", "text": "0", "font": "宋体"}, {"font_size": "18", "text": "0", "font": "宋体"}, {"font_size": "18", "text": "08", "font": "宋体"}, {"font_size": "18", "text": "5", "font": "宋体"}, {"font_size": "18", "text": "8", "font": "宋体"}, {"font_size": "18", "text": "）", "font": "楷体"}, {"font_size": "18", "text": "：春华望秋实岁物待丰成", "font": "楷体"}, {"font_size": "18", "text": "—18 ", "font": "宋体"}, {"font_size": "18", "text": "年年报点评暨 ", "font": "楷体"}, {"font_size": "18", "text": "19 ", "font": "宋体"}, {"font_size": "18", "text": "春", "font": "楷体"}, {"font_size": "18", "text": "糖见闻之二》", "font": "楷体"}, {"font_size": "18", "text": "2019-03-28", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "18", "text": "2", "font": "宋体"}, {"font_size": "18", "text": "、《五粮液（", "font": "楷体"}, {"font_size": "18", "text": "000858", "font": "宋体"}, {"font_size": "18", "text": "）", "font": "楷体"}, {"font_size": "18", "text": "—", "font": "宋体"}, {"font_size": "18", "text": "营销改革", "font": "楷体"}, {"font_size": "18", "text": "落地，普五蓄势提价》", "font": "楷体"}, {"font_size": "18", "text": "2019-03-04", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "18", "text": "3", "font": "宋体"}, {"font_size": "18", "text": "、《五粮液（", "font": "楷体"}, {"font_size": "18", "text": "000858", "font": "宋体"}, {"font_size": "18", "text": "）", "font": "楷体"}, {"font_size": "18", "text": "—", "font": "宋体"}, {"font_size": "18", "text": "来年目标", "font": "楷体"}, {"font_size": "18", "text": "积极，期待改革加速》", "font": "楷体"}, {"font_size": "18", "text": "2018-12-21", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "20", "text": "\n", "font": "Times New Roman"}, {"font_size": "21", "text": "\n", "font": "Times New Roman"}, {"font_size": "21", "text": "杨勇胜", "font": "楷体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "18", "text": "021-68407562", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "18", "text": "yangys5@cmschina.com.cn", "font": "宋体"}, {"font_size": "18", "text": " S1090514060001", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}, {"font_size": "21", "text": "欧阳予", "font": "楷体"}, {"font_size": "18", "text": "ouyangyu@cmschina.com.cn", "font": "宋体"}, {"font_size": "18", "text": " S1090519040003", "font": "宋体"}, {"font_size": "21", "text": "\n", "font": "宋体"}]}

def extract_table_1(textbox):
    textbox_content = textbox['textboxContent']
    row_list = []
    temp_list = []
    for i in range(len(textbox_content)):
        temp_list.append(textbox_content[i])
        if textbox_content[i]['text'] == '\n':
            row_list.append(temp_list)
            temp_list = []
    return row_list


def check_row(row):
    table_list = []
    cell_flag = False
    cell = []

    # print(row)
    # print("***********")

    for i in range(len(row)):
        if row[i]['text'] == '\t' or row[i]['text'] == '\n':
            table_list.append(i)

    cell.append(row[0:table_list[0]])
    for j in range(len(table_list)):
        cell.append(row[table_list[j]+1:table_list[j+1]])
        # print(cell)
        if j == len(table_list)-2:
            break
    return cell



def check_table(row_list):
    table_list = []
    tab_num = ''
    table = []
    table_flag = False

    for i in range(len(row_list)):
        if len(row_list[i]) == 2 and row_list[i][1]['text'] == '\n' and len(table) == 0:
            table.append([row_list[i]])
            table_flag = True
            continue

        if len(row_list[i]) == 2 and row_list[i][1]['text'] == '\n':
            return table

        if table_flag == True:
            cell_list = check_row(row_list[i])
            table.append(cell_list)
            continue

    return table


def coordinate_calculate(height,width,margin_top,margin_left,height_,width_,margin_top_,margin_left_):
    height_flag = False
    width_flag = False
    margin_left_flag = False
    margin_top_flag = False

    height_min = height_ - height_ * 0.1
    height_max = height_ + height_ * 0.1
    width_min = width_ - width_ * 0.1
    width_max = width_ + width_ * 0.1
    margin_top_min = margin_top_ - margin_top_ * 0.1
    margin_top_max = margin_top_ + margin_top_ * 0.1
    margin_left_min = margin_left_ - margin_left_ * 0.1
    margin_left_max = margin_left_ + margin_left_ * 0.1

    if height > height_min and height < height_max:
        height_flag = True

    if width > width_min and width < width_max:
        width_flag = True

    if margin_left > margin_left_min and margin_left < margin_left_max:
        margin_left_flag = True

    if margin_top > margin_top_min and margin_top < margin_top_max:
        margin_top_flag = True

    if height_flag and width_flag and margin_left_flag and margin_top_flag:
        return True
    else:
        return False


def locate_table_1(text_dict,height_,width_,margin_top_,margin_left_):
    for text_box in text_dict:
        textbox_shape = text_box['textboxShape']

        height = textbox_shape['height']
        width = textbox_shape['width']
        margin_top = textbox_shape['y']
        margin_left = textbox_shape['x']

        # height = float(height.replace('pt',''))
        # width = float(width.replace('pt',''))
        # margin_top = float(margin_top.replace('pt',''))
        # margin_left = float(margin_left.replace('pt',''))
        table_flag = coordinate_calculate(height,width,margin_top,margin_left,height_,width_,margin_top_,margin_left_)
        if table_flag:
            return text_box
    return []


def get_string(text):
    if text is None:
        text = ""
    if type(text) is not str:
        text = str(text)
    return text


def parse_shape(text):
    shape_dict = {}
    if len(text) == 0:
        return shape_dict
    if ";" not in text:
        print(text)
        return shape_dict
    sub_list = text.split(";")
    for line in sub_list:
        if len(line) == 0:
            continue
        if ":" not in text:
            continue
        temp_list = line.split(":")
        if len(temp_list) < 2:
            continue
        shape_dict[temp_list[0]] = temp_list[1]
    return shape_dict



def parse_report(file_path):
    # document = docx.Document(file_path)
    response = requests.get(file_path)
    with open(temp_report_path, 'wb') as file:
        file.write(response.content)
    file.close()
    f = open(temp_report_path, 'rb')
    document = docx.Document(f)
    f.close()

    children = document.element.body.iter()
    textbox_list = []
    textbox_shape_list = []
    temp_shape = ""
    for child in children:
        if child.tag.endswith('shape'):
            temp_shape = child.attrib.values()[0]
        if child.tag.endswith('textbox'):
            sub_list = []
            for item in child.iter():
                sub_list.append(item)
            if len(sub_list) > 0:
                textbox_list.append(sub_list)
                textbox_shape_list.append(temp_shape)

    print(len(textbox_list))
    # print(textbox_list)
    print(len(textbox_shape_list))

    total_dict_list = []
    for i in range(len(textbox_list)):
        textbox = textbox_list[i]
        shape = textbox_shape_list[i]
        shape_dict = parse_shape(shape)

        textbox_dict_list = []
        temp_font = "宋体"
        temp_font_size = "21"
        for child in textbox:
            # print(child)
            tag = child.tag
            # print(tag)
            text = get_string(child.text)
            # print(text)
            attrib = get_string(child.attrib)

            if tag.endswith("main}pPr") and len(textbox_dict_list) > 0:
                text = "\n"

            if tag.endswith("main}rFonts") and len(attrib) > 0:
                temp_font = child.attrib.values()[0]
            if tag.endswith("main}sz") and len(attrib) > 0:
                temp_font_size = child.attrib.values()[0]
            if tag.endswith("main}tab"):
                textbox_dict_list.append({"text": '\t',
                                          "font": temp_font,
                                          "font_size": temp_font_size})

            if (tag.endswith("main}t") or tag.endswith("main}pPr")) and len(text) > 0:
                textbox_dict_list.append({"text": text,
                                          "font": temp_font,
                                          "font_size": temp_font_size})
                temp_font = "宋体"
                temp_font_size = "21"
        if len(textbox_dict_list) > 0:
            total_dict_list.append({"textbox_content": textbox_dict_list,
                                    "textbox_shape": shape_dict})
    return total_dict_list


def check_report(file_path):
    if len(file_path) == 0:
        return None
    parse_result = parse_report(file_path)
    return parse_result


def judge_text(text):
    try:
        params = dict()
        params['content'] = text
        data = json.dumps(params).encode("utf-8")
        result = requests.post(nlp_url, data=data)
        result_json = result.json()
        return result_json

    except Exception as e:
        # logger.exception("Exception: {}".format(e))
        print(e)
        return []

# def extract_name(textbox):
#     textbox_content = textbox['textbox_content']
#     for text_dict in textbox_content:
#         text = text_dict['text']
#
#         text_json = judge_text(text)
#         text_json['data']


def fix_coordinate(docx_results):
    font_18_x = 10.5
    font_18_y = 10.5

    list_of_textbox = []
    textbox_info = []
    for docx_result in docx_results:
        # docx_dict = json.loads(docx_result)
        # print(docx_result)
        textbox_content = docx_result['textboxContent']
        textbox_shape = docx_result['textboxShape']
        textbox_width = textbox_shape['width']
        textbox_height = textbox_shape['height']
        textbox_x = textbox_shape['x']
        textbox_y = textbox_shape['y']
        # textbox_page = docx_dict['page']
        textbox_page = 1
        textbox_text = ''

        textbox_info = {"position":[textbox_x,textbox_y],
                        "page":textbox_page
                        }

        position_x = textbox_x
        position_y = textbox_y
        chars_info = []
        for run_item in textbox_content:
            font_size = run_item['fontSize']
            text = run_item['text']
            font = run_item['fontName']
            textbox_text = textbox_text + text

            if text == '\n':
                position_x = textbox_x
                position_y = position_y + font_18_y
                continue

            for w in text:
                position_x = position_x + font_18_x
                chars_dict = {'char':w,
                              'position':[position_x,position_y],
                              'font_size':font_size,
                              'font':font
                              }
                chars_info.append(chars_dict)
                if position_x+font_18_x >= textbox_x + textbox_width:
                    position_x = textbox_x
                    position_y = position_y + font_18_y

        textbox_info['chars_info'] = chars_info
        textbox_info['text'] = textbox_text

        list_of_textbox.append(textbox_info)
    print(list_of_textbox)


if __name__ == "__main__":

    width = 150.5
    height = 593.4
    y = 179.3
    x = 42.6

    width = 150.5
    height = 593.4
    y = 179.3
    x = 42.6

    file_path = 'http://10.0.0.112/group112/M00/09/5A/CgAAb17E5gzg9dW5AAGN3Z3P4vw70.docx'
    # parse_result = check_report(file_path)
    # print(parse_result)
    print("****************************************")

    url = 'http://semantic-wordocparse-service:31001/swagger-ui.html#/wordoc-parse-controller/parseZSZQUsingGET'
    url = 'http://semantic-wordocparse-service:31001/parse?wordUrl=http://10.0.0.112/group112/M00/09/5A/CgAAb17E5gzg9dW5AAGN3Z3P4vw70.docx'

    post_data = {}

    response = requests.get(url, data=json.dumps(post_data))
    # print('status: ', response)

    response = response.json()
    # print('response: ', response)

    parse_result = response['data']
    # print(parse_result)
    # text_dict_list = []
    # text_dict_list.append(text_dict)
    fix_coordinate(parse_result)
    # textbox = locate_table_1(parse_result,height,width,y,x)
    # #
    # row_list = extract_table_1(textbox)
    # print("----------------------------------")
    # table = check_table(row_list)
    # print(table)


