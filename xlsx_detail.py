# -*- coding: utf-8 -*-
import openpyxl
from handle_excel import *
import shutil
import sys
import unicodedata
import re

class xlsx_position:
    col_detial = 0          # 付款事项说明列号
    col_name = 0            # 产品名称列号
    col_spec = 0            # 规格对应的列号
    col_number = 0           # 数量列号
    col_unit_price = 0      # 单价列号
    col_total_price = 0     # 总价列号
    def __init__(self, col_detial, col_total_price):
        self.col_detial = col_detial
        self.col_total_price = col_total_price
        # 规格、数量、单价列自动生成
        self.col_name = self.col_detial + 1
        self.col_spec = self.col_name + 1
        self.col_number = self.col_spec + 1
        self.col_unit_price = self.col_number + 1
        
    def auto_insert_col(self, xlsx_handle):
        # 自动插入列
        xlsx_handle.insert_col(self.col_name, "名称") # 插入名称
        xlsx_handle.insert_col(self.col_spec, "规格") # 插入规格
        xlsx_handle.insert_col(self.col_number, "数量") # 插入数量
        xlsx_handle.insert_col(self.col_unit_price, "单价") # 插入单价

def is_chinese(string):
    """
    检查整个字符串是否包含中文
    :param string: 需要检查的字符串
    :return: bool
    """
    for ch in string:
        if u'\u4e00' <= ch <= u'\u9fff':
            return True
    return False

def go_split(s, symbol):
    """
    根据多个分隔符，全部拆分字符串,symbol 字符串集合:
    symbol = ';./+'
    """
    # 拼接正则表达式
    symbol = "[" + symbol + "]+"
    # 一次性分割字符串
    result = re.split(symbol, s)
    # 去除空字符
    return [x for x in result if x]

class detial_parse:
    # result = []             # 解析结果保持,列表字典
    '''
    付款事项解析
    '''
    def parse_detail_param(self, details):
        '''
        解析付款事项说明
        返回字典数据:解析后的详细事项(detail)， 名称(name)， 规格(spec)， 数量(number)， 单价(unit_price)， 总价(total_price)
        如果有多条记录，则返回字典数据的元组
        '''
        self.result = []             # 解析结果保持,列表字典
        # 每条解析出来后，专成英文特殊字符
        details_en = unicodedata.normalize('NFKC', details)
        detail_list = self.parse_details_spilt_cell(details_en)     # 返回值是个列表  
        for detail in detail_list:
            cell_dict = {}              # 解析出来的字典数据
            cell_list = self.parse_detail_spilt(detail)          # 返回值是个列表 ， cell_list[0]作为名称 
            spec = self.parse_spec_param(cell_list[0])     # 解析规格
            # print("规格:", spec)
            cell_dict["detail"] = detail
            cell_dict["name"] = cell_list[0]
            cell_dict["spec"] = spec
            self.result.append(cell_dict)
        return self.result
    def parse_details_spilt_cell(self, details):
        '''
        拆解一个单元格内的多个产品
        返回规格数据
        '''
        symbol = ';\n'
        detail_list = go_split(details, symbol)
        # print("detail_list:", detail_list)
        return detail_list
    def parse_detail_spilt(self, detail):
        '''
        后续规格建议与名称一起，放到括号内，以冒号分割其他字符
        解析出规格，规格与名称是一个字符串，中英文区分-冒号(:) > 空格区分 > 逗号(,)
        能解析出来规格，返回规格数据，解析失败返回None
        '''
        symbol = ':, '
        cell_list = go_split(detail, symbol)
        print("detail_split:", cell_list)
        return cell_list
    def parse_spec_param(self, name_spec):
        '''
        后续规格建议与名称一起，放到括号内，以冒号分割其他字符
        解析流程:
        1、名称中有以括号()区分规格，括号内是规格
        2、无任何规格区分，需判断非中文汉字内容，直接作为规格使用
        能解析出来规格，返回规格数据，解析失败返回None
        '''
        spec = re.findall(r'[(](.*?)[)]', name_spec)    # 获取到时，返回列表,否则返回原字符串
        # print("name_spec:", name_spec, "spec:", spec)
        if isinstance(spec, list) == True and len(spec) > 0:      # 找到规格
            return ''.join(spec)
        spec = re.sub(r'[^A-Za-z0-9\-\,\。.*]+', ' ', name_spec)      # 先把其余字符串替换成空格
        spec = re.sub(r'[ ]+', ' ', spec)   # 去除多余的空格
        # spec = re.sub("[^A-Za-z0-9\,\。]", "", name_spec)
        # print("spec param:", spec)
        if is_chinese(spec) == False:       # 解析的规格中没有中文，则认为解析成功
            return ''.join(spec)
        # 以上都没解析到，则任务解析失败
        print("规格解析失败，请使用括号或者纯英文和数字标识!:", name_spec)
        return None
    def parse_number_param(self, param):
        '''
        解析出数量
        '''
        

        
    def parse_unit_price_param(self):
        '''
        解析出单价
        '''

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage:  xlsx_detail [file.xlsx]')
        sys.exit()
    src_file = sys.argv[1]
    dst_file = "明细表_" + src_file
    print("需要处理的文件名称:", src_file, "转换后的文件名称是:", dst_file)
    shutil.copy(sys.argv[1], dst_file)

    xlsx_handle = handle_excel(dst_file)
    xlsx_pos = xlsx_position(xlsx_handle.get_col_num_from_row(1, "付款事项情况说明"), xlsx_handle.get_col_num_from_row(1, "付款金额"))
    xlsx_pos.auto_insert_col(xlsx_handle)       # 自动插入列数据

    print("get col line:", xlsx_pos.col_detial, "row max:", xlsx_handle.get_row_max(), "col max:", xlsx_handle.get_col_max())

    # 把所有行数据取出来
    all_rows = []
    for row_tmp in range(xlsx_handle.get_row_max()):
        all_rows.append(xlsx_handle.get_row_list(row_tmp+1))         # 获取一行数据

    cur_row_num = 1         # 当前行号
    for row in all_rows:
        if cur_row_num == 1:
            cur_row_num += 1
            continue

        print(row)
        detial_sxls = detial_parse()
        dst_value = row[xlsx_pos.col_detial - 1]           # detial 是Excel中的序号，从1开始
        result_list = detial_sxls.parse_detail_param(dst_value)
        if len(result_list) > 1:
             xlsx_handle.copy_row(cur_row_num, len(result_list) - 1)
        # 根据解析结果，设置单元格数据
        for num in range(len(result_list)):
            xlsx_handle.set_cell_value(cur_row_num + num, xlsx_pos.col_name, result_list[num]['name'])
            xlsx_handle.set_cell_value(cur_row_num + num, xlsx_pos.col_spec, result_list[num]['spec'])
            # xlsx_handle.set_cell_value(cur_row_num + num, xlsx_pos.col_number, result_list[num]['number'])
            # xlsx_handle.set_cell_value(cur_row_num + num, xlsx_pos.col_unit_price, result_list[num]['total_price'])

        cur_row_num += len(result_list)
        # for result in result_list:
        # print("row_num:", row_num, "dst value:", dst_value.value)
    xlsx_handle.save()