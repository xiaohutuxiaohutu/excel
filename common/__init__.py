from configparser import ConfigParser
import pymysql
import os

class DoConfig:
  def __init__(self, filepath, encoding='utf-8'):
    self.cf = ConfigParser()
    self.cf.read(filepath, encoding)

  # 获取所有的section
  def get_sections(self):
    return self.cf.sections()

  # 获取某一section下的所有option
  def get_option(self, section):
    return self.cf.options(section)

  # 获取section、option下的某一项值-str值
  def get_strValue(self, section, option):
    return self.cf.get(section, option)

  # 获取section、option下的某一项值-int值
  def get_intValue(self, section, option):
    return self.cf.getint(section, option)

  # 获取section、option下的某一项值-float值
  def get_floatValue(self, section, option):
    return self.cf.getfloat(section, option)

  # 获取section、option下的某一项值-bool值
  def get_boolValue(self, section, option):
    return self.cf.getboolean(section, option)

  def setdata(self, section, option, value):
    return self.cf.set(section, option, value)

  def get_items(self, item_name):
    return self.cf.items(item_name)

curDir = os.getcwd()  # 获取当前文件路径
config_path = curDir[:curDir.find("excel\\") + len("excel\\")]+'common\\config.ini'
# config_path = 'config.ini'


def get_mysql_conn():
  cf = DoConfig(config_path)
  user_name = cf.get_strValue('mysql_config', 'user')
  password = cf.get_strValue('mysql_config', 'password')
  host_value = cf.get_strValue('mysql_config', 'host')
  host = 'localhost' if host_value is None else host_value
  port_value = cf.get_intValue('mysql_config', 'port')
  port = 3306 if port_value is None else port_value

  database = cf.get_strValue('mysql_config', 'database')

  charset_value = cf.get_strValue('mysql_config', 'charset')
  charset ='utf8' if charset_value is None else charset_value

  conn = pymysql.connect(host=host, port=port, user=user_name, passwd=password, db=database,charset=charset)
  return conn


if __name__ == '__main__':
  get_mysql_conn()
# cf = DoConfig('demo.conf')
# res = cf.get_sections()
# print(res)
# res = cf.get_option('db')
# print(res)
# res = cf.get_strValue('db', 'db_name')
# print(res)
# res = cf.get_intValue('db', 'db_port')
# print(res)
# res = cf.get_floatValue('user_info', 'salary')
# print(res)
# res = cf.get_boolValue('db', 'is')
# print(res)
#
# cf.setdata('db', 'db_port', '3306')
# res = cf.get_strValue('db', 'db_port')
# print(res)
