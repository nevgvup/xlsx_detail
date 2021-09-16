# -*- coding: utf-8 -*-
import openpyxl
from handle_excel import *
import shutil
import sys
import unicodedata
import re

class xlsx_position:
    col_spec = 0            # 规格对应的列号
    col_detial = 0          # 付款事项说明列号
    col_count = 0           # 数量列号
    col_unit_price = 0      # 单价列号
    col_total_price = 0     # 总价列号
    def __init__(self, col_detial, col_total_price):
        self.col_detial = col_detial
        self.col_total_price = col_total_price
        # 规格、数量、单价列自动生成
        self.col_spec = self.col_detial + 1
        self.col_count = self.col_spec + 1
        self.col_unit_price = self.col_count + 1
        
    def auto_insert_col(self, xlsx_handle):
        # 自动插入列
        xlsx_handle.insert_col(self.col_spec, "规格") # 插入规格
        xlsx_handle.insert_col(self.col_count, "数量") # 插入数量
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

class detial_parse:
    '''
    付款事项解析
    '''
    def parse_detail_param(self, details):
        '''
        解析付款事项说明
        返回字典数据:解析后的详细事项(detail)， 名称(name)， 规格(spec)， 数量(number)， 单价(unit_price)， 总价(total_price)
        如果有多条记录，则返回字典数据的元组
        '''
        # 每条解析出来后，专成英文特殊字符
        details_en = unicodedata.normalize('NFKC', details)

    def parse_detail_spilt(self, detail):
        '''
        后续规格建议与名称一起，放到括号内，以冒号分割其他字符
        解析出规格，规格与名称是一个字符串，中英文区分-冒号(:) > 空格区分 > 逗号(,)
        能解析出来规格，返回规格数据，解析失败返回None
        '''
        spilt_status = False
        # split_detail = detail.split("：")
        # if len(split_detail) > 1:     # 找到分割符中文冒号：
        #     spilt_status = True

        if spilt_status == False:
            split_detail = detail.split(":")   # 找到分割符英文冒号:
            if len(split_detail) > 1:
                spilt_status = True

        if spilt_status == False:
            split_detail = detail.split(" ")
            if len(split_detail) > 1:     # 找到分割符空格
                spilt_status = True

        # if spilt_status == False:
        #     split_detail = detail.split("，")
        #     if len(split_detail) > 1:     # 找到分割符中文逗号
        #         spilt_status = True

        if spilt_status == False:
            split_detail = detail.split(",")
            if len(split_detail) > 1:     # 找到分割符英文逗号
                spilt_status = True

        if spilt_status== False:
            print("错误!!未找到名称与规格区分，请使用冒号、逗号、空格区分:", detail)
            return None

    def parse_spec_param(self, detail):
        '''
        后续规格建议与名称一起，放到括号内，以冒号分割其他字符
        解析流程:
        1、名称中有以括号()区分规格，括号内是规格
        2、无任何规格区分，需判断非中文汉字内容，直接作为规格使用
        能解析出来规格，返回规格数据，解析失败返回None
        '''
        
        detail_left = split_detail[0]

        spec = re.findall(r'[(](.*?)[)]', name_spec)    # 获取到时，返回列表,否则返回原字符串
        if isinstance(spec, list) == True:      # 找到规格
            return spec
        spec = re.sub("[^A-Za-z0-9\,\。]", "", name_spec)
        if is_chinese(spec) == False:       # 解析的规格中没有中文，则认为解析成功
            return spec
        #以上都没解析到，着任务解析失败
        print("规格解析失败，请使用括号或者纯英文和数字标识!:", name_spec)
        return None
    def parse_count_param(self):
        '''
        解析出数量
        '''
    def parse_unit_price_param(self):
        '''
        解析出单价
        '''

if __name__ == '__main__':
    if len(sys.argv)<2:
        print('Usage:  xlsx_detail [file.xlsx]')
        sys.exit()
    src_file = sys.argv[1]
    dst_file = "明细表_" + src_file
    print("需要处理的文件名称:", src_file, "转换后的文件名称是:", dst_file)
    shutil.copy(sys.argv[1], dst_file);

    xlsx_handle = handle_excel(dst_file)
    xlsx_pos = xlsx_position(xlsx_handle.get_col_num_from_row(1, "付款事项情况说明"), xlsx_handle.get_col_num_from_row(1, "付款金额"))
    xlsx_pos.auto_insert_col(xlsx_handle)       # 自动插入列数据

    print("get col line:", xlsx_pos.col_detial, "col max:", xlsx_handle.get_col_max())

   
    for row_num in range(xlsx_handle.get_col_max()):
        row_num += 1
        dst_value = xlsx_handle.get_cell_value(row_num, xlsx_pos.col_detial)
        print("row_num:", row_num, "dst value:", dst_value.value)




