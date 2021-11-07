import configparser
import os

class ReadConfig:
    """定义一个读取配置文件的类"""

    def __init__(self, filepath=None):
        if filepath:
            configpath = filepath
        else:
            root_dir = os.path.abspath('.')
            configpath = os.path.join(root_dir, "config.ini")
        self.cf = configparser.ConfigParser()
        self.cf.read(configpath, encoding='utf-8')
    # 获取单位键值
    def get_unit(self):
        value = ''
        if self.cf.has_option('Config', 'unit'): 
            value = self.cf.get("Config", "unit")
        return value
    # 获取分句键值
    def get_paragraph(self):
        value = ''
        if self.cf.has_option('Config', 'paragraph'): 
            value = self.cf.get("Config", "paragraph")
        return value
    # 获取多个产品分割字符
    def get_category(self):
        value = ''
        if self.cf.has_option('Config', 'category'): 
            value = self.cf.get("Config", "category")
        return value
