# -*- coding: utf-8 -*-
import openpyxl
from handle_excel import *
from readconfig import *
import shutil
import sys
import unicodedata
import re

class xlsx_position:
    col_detial = 0          # 付款事项说明列号
    col_name = 0            # 产品名称列号
    col_spec = 0            # 规格对应的列号
    col_number = 0          # 数量列号
    col_unit = 0            # 单位列号
    col_unit_price = 0      # 单价列号
    col_total_price = 0     # 总价列号
    col_total_pay = 0       # 付款总价
    def __init__(self, col_detial, col_total_pay):
        self.col_detial = col_detial
        self.col_total_pay = col_total_pay
        # 规格、数量、单价列自动生成
        self.col_name = self.col_detial + 1
        self.col_spec = self.col_name + 1
        self.col_number = self.col_spec + 1
        self.col_unit = self.col_number + 1
        self.col_unit_price = self.col_unit + 1
        self.col_total_price = self.col_unit_price + 1
    def auto_insert_col(self, xlsx_handle):
        # 自动插入列
        xlsx_handle.insert_col(self.col_name, "名称") # 插入名称
        xlsx_handle.insert_col(self.col_spec, "规格") # 插入规格
        xlsx_handle.insert_col(self.col_number, "数量") # 插入数量
        xlsx_handle.insert_col(self.col_unit, "单位") # 插入单位
        xlsx_handle.insert_col(self.col_unit_price, "单价") # 插入单价
        xlsx_handle.insert_col(self.col_total_price, "总价") # 插入总价
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

def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass
 
    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass
 
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
        返回字典数据:解析后的详细事项(detail)， 名称(name)， 规格(spec)， 数量(number)， 单位(unit)， 单价(unit_price)， 总价(total_price)
        如果有多条记录，则返回字典数据的元组
        '''
        self.result = []             # 解析结果保持,列表字典
        # 每条解析出来后，专成英文特殊字符
        details_en = unicodedata.normalize('NFKC', details)
        detail_list = self.parse_details_spilt_cell(details_en)     # 返回值是个列表  
        for detail in detail_list:
            cell_dict = {}              # 解析出来的字典数据
            cell_list = self.parse_detail_spilt(detail)          # 返回值是个列表 ， cell_list[0]作为名称 
            # 判断cell_list是否为空
            if (len(cell_list) == 0):
                continue

            spec, name = self.parse_name_spec_param(cell_list[0])     # 解析规格
            for i in range(1, len(cell_list)):
                if "number" not in  cell_dict:
                    parse_num, unit = self.parse_number_param(cell_list[i])       # 解析数量
                    if len(parse_num) > 0 and is_number(parse_num):
                        cell_dict["number"] = float(parse_num)
                        cell_dict["unit"] = unit
                if "unit_price" not in  cell_dict:
                    unit_price = self.parse_unit_price_param(cell_list[i])       # 解析单价
                    if len(unit_price) > 0 and is_number(unit_price):
                        cell_dict["unit_price"] = float(unit_price)
            # print("规格:", spec)
            cell_dict["detail"] = detail
            cell_dict["name"] = name
            cell_dict["spec"] = spec
            # 总价计算,需要存在数量和单价
            if "number" in  cell_dict and "unit_price" in  cell_dict:
                cell_dict["total_price"] = (float(cell_dict["number"]) * float(cell_dict["unit_price"]))
            self.result.append(cell_dict)
        return self.result
    def parse_details_spilt_cell(self, details):
        '''
        拆解一个单元格内的多个产品
        返回规格数据
        '''
        cfg = ReadConfig()
        cfg_symbol = cfg.get_category()      # 获取不同购买品类
        if len(cfg_symbol) <= 0:
            cfg_symbol = ';\n'

        detail_list = go_split(details, cfg_symbol)
        # print("detail_list:", detail_list)
        return detail_list
    def parse_detail_spilt(self, detail):
        '''
        后续规格建议与名称一起，放到括号内，以冒号分割其他字符
        解析出规格，规格与名称是一个字符串，中英文区分-冒号(:) > 空格区分 > 逗号(,)
        能解析出来规格，返回规格数据，解析失败返回None
        '''
        cfg = ReadConfig()
        cfg_symbol = cfg.get_paragraph()      # 获取分段
        if len(cfg_symbol) <= 0:
            cfg_symbol = ':, '

        cell_list = go_split(detail, cfg_symbol)
        print("detail_split:", cell_list)
        return cell_list
    def parse_name_spec_param(self, name_spec):
        '''
        后续规格建议与名称一起，放到括号内，以冒号分割其他字符
        解析流程:
        1、名称中有以括号()区分规格，括号内是规格
        2、无任何规格区分，需判断非中文汉字内容，直接作为规格使用
        能解析出来规格，返回规格数据，解析失败返回None
        '''
        # 先去除名称中的序号1、 2、 3、
        # name_spec_tmp = re.findall(r'.*、', name_spec)    # 获取到时，返回列表,否则返回原字符串
        name_spec_tmp = name_spec
        name_spec_obj = re.search(r'(.*)、(.*)', name_spec)
        if name_spec_obj:
            print("可能存在序号!")
            # 判断group1是否是数字
            if is_number(name_spec_obj.group(1)):
                print("当前是序号")
                name_spec_tmp = name_spec_obj.group(2)
            # print("name_spec_obj.group():", name_spec_obj.group())
            # print("name_spec_obj.group(1):", name_spec_obj.group(1))
            # print("name_spec_obj.group(2):", name_spec_obj.group(2))

        print("name22222 去掉 序号 spec:", name_spec_tmp)
        name_spec_m = re.match(r'(.*?)\((.*?)\)(.*?)', name_spec_tmp)
        if name_spec_m:      # 找到规格,group2是规格
            print("group(1):", name_spec_m.group(1), "group(2):", name_spec_m.group(2))
            return ''.join(name_spec_m.group(2)), ''.join(name_spec_m.group(1))
        spec = re.sub(r'[^A-Za-z0-9\-\,\。\.\*]+', ' ', name_spec_tmp)      # 先把其余字符串替换成空格
        spec = re.sub(r'[ ]+', ' ', spec)   # 去除多余的空格
        # spec = re.sub("[^A-Za-z0-9\,\。]", "", name_spec_tmp)
        # print("spec param:", spec)
        if is_chinese(spec) == False:       # 解析的规格中没有中文，则认为解析成功
            return ''.join(spec), name_spec_tmp
        # 以上都没解析到，则任务解析失败
        print("规格解析失败，请使用括号或者纯英文和数字标识!:", name_spec_tmp)
        return None
    def parse_number_param(self, param):
        '''
        解析出数量
        '''
        cfg = ReadConfig()
        cfg_unit = cfg.get_unit()      # 获取单位配置
        if len(cfg_unit) <= 0:
            cfg_unit = r'.*个|.*套|.*只|.*g|.*kg|.*支|.*张|.*箱|.*桶|.*包|.*卷|.*根|.*米|.*升|.*瓶|.*袋'
        # print("cfg_unit:", cfg_unit)
        numer_tmp = re.findall(cfg_unit, param)
        print("parse number param:", param, "numer:", numer_tmp)
        numer_tmp = ''.join(numer_tmp)
        numer = re.sub(r'[^0-9.]', '', numer_tmp)    # 获取数值0-9.
        unit = re.findall(r'[^0-9.]', numer_tmp)    # 单位
        return numer, ''.join(unit)
    def parse_unit_price_param(self, param):
        '''
        解析出单价
        '''
        price = re.findall(r'.*元/|单价.*元|\*.*元', param)
        print("param:", param, "price:", ''.join(price))
        price = ''.join(price)
        price = re.sub(r'[^0-9.]', '', price)    # 获取数值0-9.

        return ''.join(price)


# 处理xlsx文件
def detail_xlsx(dst_file):
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
            if "name" in  result_list[num]:
                xlsx_handle.set_cell_value(cur_row_num + num, xlsx_pos.col_name, result_list[num]['name'])
            if "spec" in  result_list[num]:
                xlsx_handle.set_cell_value(cur_row_num + num, xlsx_pos.col_spec, result_list[num]['spec'])
            if "number" in  result_list[num]:
                xlsx_handle.set_cell_value(cur_row_num + num, xlsx_pos.col_number, result_list[num]['number'])
            if "unit" in  result_list[num]:
                xlsx_handle.set_cell_value(cur_row_num + num, xlsx_pos.col_unit, result_list[num]['unit'])
            if "unit_price" in  result_list[num]:
                xlsx_handle.set_cell_value(cur_row_num + num, xlsx_pos.col_unit_price, result_list[num]['unit_price'])
            if "total_price" in  result_list[num]:
                xlsx_handle.set_cell_value(cur_row_num + num, xlsx_pos.col_total_price, result_list[num]['total_price'])
        cur_row_num += len(result_list)
        # for result in result_list:
        # print("row_num:", row_num, "dst value:", dst_value.value)
    xlsx_handle.save()

# 查找某个文件夹及其子文件夹下指定后缀名的所有文件
def findAllFilesWithSpecifiedSuffix(target_dir, target_suffix="xlsx"):
    find_res = []
    target_suffix_dot = "." + target_suffix
    walk_generator = os.walk(target_dir)
    for root_path, dirs, files in walk_generator:
        if len(files) < 1:
            continue
        for file in files:
            file_name, suffix_name = os.path.splitext(file)
            if suffix_name == target_suffix_dot:
                find_res.append(os.path.join(root_path, file))
    return find_res

def mkdir(path):
    # 去除首位空格
    path=path.strip()
    # 去除尾部 \ 符号
    path=path.rstrip("\\")

    # 判断路径是否存在
    # 存在     True
    # 不存在   False
    isExists=os.path.exists(path)

    # 判断结果
    if not isExists:
        # 如果不存在则创建目录
        # 创建目录操作函数
        os.makedirs(path)
        return True
    else:
        # 如果目录存在则不创建，并提示目录已存在
        return False

if __name__ == '__main__':
    root_dir = os.path.abspath('.')
    dst_dir = os.path.join(root_dir, "结果输出")
    file_list = findAllFilesWithSpecifiedSuffix(root_dir)
    if len(file_list) <= 0:
        print("未找到xlsx文件")
        sys.exit()
    mkdir(dst_dir)
    print("file_list:", file_list)
    for src_file in file_list:
        file_name = os.path.basename(src_file)  # 返回文件名
        dst_file = os.path.join(dst_dir, file_name)
        print("dst file:", dst_file)
        shutil.copy(src_file, dst_file)
        detail_xlsx(dst_file)
