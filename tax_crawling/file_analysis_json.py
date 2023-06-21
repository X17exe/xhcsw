from selenium import webdriver
from time import sleep
from selenium.webdriver import ChromeOptions
from selenium.webdriver.support.ui import WebDriverWait
import pyautogui
import pika
import random
import os
import shutil
import requests
import yaml
import yagmail
import sys
import win32api
import pdfplumber
import re
import json
import copy
import datetime
import win32gui
import win32con
import win32com.client

def read_yaml():
    """
    读取配置文件
    """
    try:
        with open('D:/workspace/xht_sw/tax_crawling/configure.yml', encoding="utf-8") as y:
        # with open('C:/dist/monitor_downloadtax/configure.yml', encoding="utf-8") as y:
            yaml_data = yaml.safe_load(y)
        return yaml_data
    except BaseException as t:
        pass


def taxdetail_regular_extraction(*args):
    """
    读取税单表格详细内容后，传入生成详情信息字典
    """
    tax_detailed_dict = {}
    # tax_detailed_dict['税单生成日期'] = ' ,'.join(re.findall(r"(.*?)    1/", *args))
    tax_detailed_dict['报关单号'] = ' ,'.join(re.findall(r"报关单号： (.*?) 税单序号", *args))
    tax_detailed_dict['税单序号'] = ' ,'.join(re.findall(r"税单序号： (.*?) 税种：", *args))
    tax_detailed_dict['税种'] = ' ,'.join(re.findall(r"税种： (.*?) 现场税单序：", *args))
    tax_detailed_dict['现场税单序'] = ' ,'.join(re.findall(r"现场税单序： (.*?) 申报口岸：", *args))
    tax_detailed_dict['申报口岸'] = ' ,'.join(re.findall(r"申报口岸： (.*?) 进出口岸：", *args))
    tax_detailed_dict['进出口岸'] = ' ,'.join(re.findall(r"进出口岸： (.*?) 收发货单位：", *args))
    tax_detailed_dict['收发货单位'] = ' ,'.join(re.findall(r"收发货单位： (.*?) 消费使用单位：", *args))
    tax_detailed_dict['消费使用单位'] = ' ,'.join(re.findall(r"消费使用单位：(.*?) 申报单位：", *args))
    tax_detailed_dict['申报单位'] = ' ,'.join(re.findall(r"申报单位： (.*?) 提单号：", *args))
    tax_detailed_dict['提单号'] = ' ,'.join(re.findall(r"提单号： (.*?) 运输工具号：", *args))
    tax_detailed_dict['运输工具号'] = ' ,'.join(re.findall(r"运输工具号： (.*?) 合同号：", *args))
    tax_detailed_dict['合同号'] = ' ,'.join(re.findall(r"合同号： (.*?) 监管方式：", *args))
    tax_detailed_dict['监管方式'] = ' ,'.join(re.findall(r"监管方式： (.*?) 征免性质：", *args))
    tax_detailed_dict['征免性质'] = ' ,'.join(re.findall(r"征免性质： (.*?) 进/出口日期：", *args))
    tax_detailed_dict['进/出口日期'] = ' ,'.join(re.findall(r"进/出口日期：  (.*?) 进出口标志：", *args))
    tax_detailed_dict['进出口标志'] = ' ,'.join(re.findall(r"进出口标志： (.*?) 退补税标志：", *args))
    tax_detailed_dict['退补税标志'] = ' ,'.join(re.findall(r"退补税标志： (.*?) 滞报滞纳标志：", *args))
    tax_detailed_dict['滞报滞纳标志'] = ' ,'.join(re.findall(r"滞报滞纳标志：(.*?) 税款金额：", *args))
    tax_detailed_dict['税款金额'] = ' ,'.join(re.findall(r"税款金额： (.*?) 税款金额大写：", *args))
    tax_detailed_dict['税款金额大写'] = ' ,'.join(re.findall(r"税款金额大写：(.*?) 征税操作员：", *args))
    tax_detailed_dict['征税操作员'] = ' ,'.join(re.findall(r"征税操作员： (.*?) 税单开征日期：", *args))
    tax_detailed_dict['税单开征日期'] = ' ,'.join(re.findall(r"税单开征日期：(.*?) 缴款期限：", *args))
    tax_detailed_dict['缴款期限'] = ' ,'.join(re.findall(r"缴款期限： (.*?) 收入机关：", *args))
    tax_detailed_dict['收入机关'] = ' ,'.join(re.findall(r"收入机关： (.*?) 收入系统：", *args))
    tax_detailed_dict['收入系统'] = ' ,'.join(re.findall(r"收入系统： (.*?) 预算级次：", *args))
    tax_detailed_dict['预算级次'] = ' ,'.join(re.findall(r"预算级次： (.*?) 预算科目名称：", *args))
    tax_detailed_dict['预算科目名称'] = ' ,'.join(re.findall(r"预算科目名称： (.*?) 收款国库：", *args))
    tax_detailed_dict['收款国库'] = ' ,'.join(re.findall(r"收款国库： (.*?) 税费单货物信息", *args))

    return tax_detailed_dict


def goodsdetail_regular_extraction(taxno, goodsname, quantity, company, currency, taxconversionrate, dutiableprice,
                                   valoremrate, specificrate, taxamount):
    """
    读取税单货物信息后，传入生成货物信息字典
    """
    goods_detailed_dict = {}
    goods_detailed_dict['税号'] = ''.join(taxno)
    goods_detailed_dict['货名'] = ''.join(goodsname)
    goods_detailed_dict['数量'] = ''.join(quantity)
    goods_detailed_dict['单位'] = ''.join(company)
    goods_detailed_dict['币制'] = ''.join(currency)
    goods_detailed_dict['外汇折算率'] = ''.join(taxconversionrate)
    goods_detailed_dict['完税价格'] = ''.join(dutiableprice)
    goods_detailed_dict['从价税率'] = ''.join(valoremrate)
    goods_detailed_dict['从量税率'] = ''.join(specificrate)
    goods_detailed_dict['税额'] = ''.join(taxamount)

    return goods_detailed_dict


# def analysis_taxfile(file):
#     """
#     解析抓取的税单，并转为json对象
#     """
#     allgoods_list = []
#     taxfilenumber = 0
#     with pdfplumber.open(file) as t:
#         for page_content in t.pages:
#             allfile_content = page_content.extract_text().replace('\n', ' ')
#
#             #  提取税费信息,转为字典
#             if '税费单详细信息' in allfile_content:
#                 if int(taxfilenumber) == 0:
#                     taxbill_detail_dict1 = taxdetail_regular_extraction(allfile_content)
#                 elif int(taxfilenumber) == 1:
#                     taxbill_detail_dict2 = taxdetail_regular_extraction(allfile_content)
#                 taxfilenumber += 1
#
#             #  提取货物信息,转为列表
#             if '税费单货物信息' in allfile_content:
#                 detail = re.findall(r"折算率 税率 税率 (.*?)$", allfile_content)
#                 strdetail = ''.join(detail)
#                 str_detail = strdetail + ' '
#                 listdetail = re.findall(r"(.*?) ", str_detail)
#
#                 #  拆分列表，拆分后一个子列表为一行货物明细数据
#                 cd = 10
#                 if len(listdetail) > cd:
#                     for i in range(int(len(listdetail) / cd)):
#                         cut_a = listdetail[cd * i:cd * (i + 1)]
#                         allgoods_list.append(cut_a)
#
#                     last_data = listdetail[int(len(listdetail) / cd) * cd:]
#                     if last_data:
#                         allgoods_list.append(last_data)
#                 else:
#                     allgoods_list.append(listdetail)
#     t.close()
#
#     #  解包拆分后的列表，转为字典后再转json
#     if int(taxfilenumber) == 1:
#         goods_detail_list = []
#         for singe_goods in allgoods_list:
#             singetax_goodsdict = goodsdetail_regular_extraction(*singe_goods)
#             goods_detail_list.append(singetax_goodsdict)
#         taxbill_detail_dict1['税费单货物信息'] = goods_detail_list
#         taxbill_detail_json = json.dumps(taxbill_detail_dict1, ensure_ascii=False)
#
#         return str([taxbill_detail_json]).replace("'", ""), taxfilenumber
#
#     #  taxfilenumber = 2时，证明报关单下有两个文件（增值税和关税）
#     elif int(taxfilenumber) == 2:
#         goods_detail_list1 = []
#         goods_detail_list2 = []
#         #  等分增值税、关税货物明细
#         singe_taxfile_goodsdetailnumber = int(len(allgoods_list)) / int(taxfilenumber)
#         #  列表切片
#         goodsdetaillist1 = allgoods_list[:int(singe_taxfile_goodsdetailnumber)]
#         goodsdetaillist2 = allgoods_list[int(singe_taxfile_goodsdetailnumber):]
#
#         for singe_goods in goodsdetaillist1:
#             singetax_goodsdict1 = goodsdetail_regular_extraction(*singe_goods)
#             goods_detail_list1.append(singetax_goodsdict1)
#         taxbill_detail_dict1['税费单货物信息'] = goods_detail_list1
#         taxbill_detail_json1 = json.dumps(taxbill_detail_dict1, ensure_ascii=False)
#
#         for singe_goods in goodsdetaillist2:
#             singetax_goodsdict2 = goodsdetail_regular_extraction(*singe_goods)
#             goods_detail_list2.append(singetax_goodsdict2)
#         taxbill_detail_dict2['税费单货物信息'] = goods_detail_list2
#         taxbill_detail_json2 = json.dumps(taxbill_detail_dict2, ensure_ascii=False)
#
#         return str([taxbill_detail_json1, taxbill_detail_json2]).replace("'", ""), taxfilenumber
#
#
# if __name__ == '__main__':
#     file = 'C:/taxefile/20221011174751.pdf'
#     for i in range(5):
#         print(analysis_taxfile(file))


# year = datetime.datetime.today().year
# month = datetime.datetime.today().month
# day = datetime.datetime.today().day
# # date = f"{year}年{month}月{day}日"
# date = "2022年11月24日"
# cd = 10
#
# all_content_list = []
# file = 'C:/taxefile/20221124183626.pdf'
# with pdfplumber.open(file) as f:
#     for page_content in f.pages:
#         all_file_content = page_content.extract_text().replace('\n', ' ')
#         all_content_list.append(all_file_content)
#     f.close()
#
# tax_category_num = 0
# for tax_num in all_content_list:
#     if "税费单详细信息" in tax_num:
#         tax_category_num += 1
#
# if tax_category_num == 1:
#     str_all_content1 = str(all_content_list)
#     remove_text1 = "', '%s    2/2 税费单货物信息  外汇 从价 从量 税号 货名 数量 单位  币制 完税价格 税额 折算率 税率 税率" % date
#     remove_text2 = "', '%s    2/3 税费单货物信息  外汇 从价 从量 税号 货名 数量 单位  币制 完税价格 税额 折算率 税率 税率" % date
#     remove_text3 = "', '%s    3/3 税费单货物信息  外汇 从价 从量 税号 货名 数量 单位  币制 完税价格 税额 折算率 税率 税率" % date
#     str_all_content2 = str_all_content1.replace("%s" % remove_text1, "")
#     str_all_content3 = str_all_content2.replace("%s" % remove_text2, "")
#     str_all_content4 = str_all_content3.replace("%s" % remove_text3, "")
#     str_detail1 = re.findall(r"折算率 税率 税率 (.*?)']", str_all_content4)
#     str_detail = ''.join(str_detail1)
#     str_detail = str_detail + " "
#     list_detail = re.findall(r"(.*?) ", str_detail)
#
#
# if tax_category_num == 2:
#     first_tax_category_content_list = []  # 第一个税种所有的内容
#     second_tax_category_content_list = []  # 第二个税种所有的内容
#     first_all_goods_list = []  # 第一个税种的明细
#     second_all_goods_list = []  # 第二个税种的明细
#
#     if "进口关税" in all_content_list[0]:
#         with pdfplumber.open(file) as x:
#             for page_content in x.pages:
#                 first_file_content = page_content.extract_text().replace('\n', ' ')
#                 if "进口增值税" in first_file_content:
#                     break
#                 first_tax_category_content_list.append(first_file_content)
#         x.close()
#     else:
#         with pdfplumber.open(file) as x:
#             for page_content in x.pages:
#                 first_file_content = page_content.extract_text().replace('\n', ' ')
#                 if "进口关税" in first_file_content:
#                     break
#                 first_tax_category_content_list.append(first_file_content)
#         x.close()
#
#     for first_tax_category_content in all_content_list:
#         if first_tax_category_content not in first_tax_category_content_list:
#             second_tax_category_content_list.append(first_tax_category_content)
#
#     print(first_tax_category_content_list)
#     print(second_tax_category_content_list)
#
#     str_first_tax_category_content1 = str(first_tax_category_content_list)
#     remove_text1 = "', '%s    2/2 税费单货物信息  外汇 从价 从量 税号 货名 数量 单位  币制 完税价格 税额 折算率 税率 税率" % date
#     remove_text2 = "', '%s    2/3 税费单货物信息  外汇 从价 从量 税号 货名 数量 单位  币制 完税价格 税额 折算率 税率 税率" % date
#     remove_text3 = "', '%s    3/3 税费单货物信息  外汇 从价 从量 税号 货名 数量 单位  币制 完税价格 税额 折算率 税率 税率" % date
#     str_first_tax_category_content2 = str_first_tax_category_content1.replace("%s" % remove_text1, "")
#     str_first_tax_category_content3 = str_first_tax_category_content2.replace("%s" % remove_text2, "")
#     str_first_tax_category_content4 = str_first_tax_category_content3.replace("%s" % remove_text3, "")
#     str_first_detail1 = re.findall(r"折算率 税率 税率 (.*?)']", str_first_tax_category_content4)
#     str_first_detail = ''.join(str_first_detail1)
#     str_splicing_detail1 = str_first_detail + " "
#     list_first_detail = re.findall(r"(.*?) ", str_splicing_detail1)
#
#     if len(list_first_detail) > cd:
#         for i in range(int(len(list_first_detail) / cd)):
#             cut_a = list_first_detail[cd * i:cd * (i + 1)]
#             first_all_goods_list.append(cut_a)
#
#         first_data = list_first_detail[int(len(list_first_detail) / cd) * cd:]
#         if first_data:
#             first_all_goods_list.append(first_data)
#     else:
#         first_all_goods_list.append(list_first_detail)
#
#     print(first_all_goods_list)
#
#     str_second_tax_category_content1 = str(second_tax_category_content_list)
#     remove_text1 = "', '%s    2/2 税费单货物信息  外汇 从价 从量 税号 货名 数量 单位  币制 完税价格 税额 折算率 税率 税率" % date
#     remove_text2 = "', '%s    2/3 税费单货物信息  外汇 从价 从量 税号 货名 数量 单位  币制 完税价格 税额 折算率 税率 税率" % date
#     remove_text3 = "', '%s    3/3 税费单货物信息  外汇 从价 从量 税号 货名 数量 单位  币制 完税价格 税额 折算率 税率 税率" % date
#     str_second_tax_category_content2 = str_second_tax_category_content1.replace("%s" % remove_text1, "")
#     str_second_tax_category_content3 = str_second_tax_category_content2.replace("%s" % remove_text2, "")
#     str_second_tax_category_content4 = str_second_tax_category_content3.replace("%s" % remove_text3, "")
#     str_second_detail1 = re.findall(r"折算率 税率 税率 (.*?)']", str_second_tax_category_content4)
#     str_second_detail = ''.join(str_second_detail1)
#     str_splicing_detail2 = str_second_detail + " "
#     list_second_detail = re.findall(r"(.*?) ", str_splicing_detail2)
#
#     if len(list_second_detail) > cd:
#         for i in range(int(len(list_second_detail) / cd)):
#             cut_b = list_second_detail[cd * i:cd * (i + 1)]
#             second_all_goods_list.append(cut_b)
#
#         second_data = list_second_detail[int(len(list_second_detail) / cd) * cd:]
#         if second_data:
#             second_all_goods_list.append(second_data)
#     else:
#         second_all_goods_list.append(list_second_detail)
#     print(second_all_goods_list)


# name = "I20220000954907013.pdf"
# file_name = "C:\\taxefile\\I20220000954907013.pdf"
# name = "20221130094424.pdf"
# file_name = "C:\\taxefile\\20221130094424.pdf"
# #  接口参数
# url = "https://api.delchannel.com/customs/hgBill/declare/uploadBillMultipart/532020221200201658/2"
# # url = "http://192.168.2.87:9010/import/operationCustomsBillReportInfo/declare/uploadBillMultipart/{customsNo}/{fileType}"
# content = open(file_name, 'rb')
# files = {'file': (name, content, 'pdf')}
# data = {'Content-Disposition': 'form-data', 'Content-Type': 'application/pdf'}
# r = requests.post(url=url, data=data, files=files)
# print(r.text)
# print(r.json()['code'])
# print(r.json()['msg'])

# test_list = []
# def goods_detail_dict(tax_no, cus_declaration_no, goods_name, tax_category):
#     detail_dict = {}
#
#     detail_dict["clTaxRate"] = float(0)
#     detail_dict["ctaxRate"] = float(0.13)
#     detail_dict["customsCode"] = tax_no
#     detail_dict["customsNo"] = cus_declaration_no
#     detail_dict["itemNo"] = int(1)
#     detail_dict["productCustomsName"] = goods_name
#     detail_dict["taxMoney"] = float(4547.64)
#     detail_dict["taxType"] = tax_category
#
#     return detail_dict
#
# test_list.append(goods_detail_dict('8542399000', '530120231010002277', '集成电路', "L"))
# json = json.dumps(test_list, ensure_ascii=False)
# print(type(json))
#
# data = [{"clTaxRate": 0.0, "ctaxRate": 0.13, "customsCode": "8542399000", "customsNo": "530120231010002277", "itemNo": 1, "productCustomsName": "集成电路", "taxMoney": 4547.64, "taxType": "L"}]
# json_data = json.dumps(data, ensure_ascii=False)
# print(type(data))
# header = {'Content-Type': 'application/json'}
# url = "https://api.delchannel.com/customs/hgBill/declare/uploadTaxJson"


# def upload_taxjson(detail_data):
#     """
#     上传税单json对象
#     """
#     url = 'https://api.delchannel.com/customs/hgBill/declare/uploadTaxJson'
#     header = {'Content-Type': 'application/json;charset=UTF-8'}
#     r = requests.post(url=url, headers=header, data=detail_data.encode("utf-8"))
#     code = int(r.json()['code'])
#     msg = r.json()['msg']
#     print(detail_data)
#     if code == 200:
#         print("数据上传成功！")
#     else:
#         print("数据上传失败,原因：%s" % (msg))
#
#
# upload_taxjson(detail_data=json)




def create_debugwindow():
    # 创建chrome浏览器调试窗口,创建窗口前先杀掉可能存在的debug窗口
    cmd = read_yaml()['localfile']['chromedebugwindow']
    win32api.ShellExecute(0, 'open', cmd, '', '', 1)
    sleep(3)

    option = ChromeOptions()
    option.add_argument('--no-sandbox')
    # prefs = {'profile.default_content_settings.popups': 0,
    #          'download.default_directory': 'C:\\taxefile'}
    # option.add_experimental_option('prefs', prefs)
    option.add_experimental_option("debuggerAddress", "127.0.0.1:9000")  # 打开已开启调试窗口
    browser = webdriver.Chrome(options=option)
    login_taxpage(browser)


def top_windows():
    """
    在windows置顶chrome_debug浏览器窗口
    """
    def get_all_hwnd(hwnd, mouse):
        if (win32gui.IsWindow(hwnd) and
                win32gui.IsWindowEnabled(hwnd) and
                win32gui.IsWindowVisible(hwnd)):
            hwnd_map.update({hwnd: win32gui.GetWindowText(hwnd)})

    hwnd_map = {}
    win32gui.EnumWindows(get_all_hwnd, 0)

    for h, t in hwnd_map.items():
        if t:
            if t == '中国国际贸易单一窗口 - Google Chrome' or t == '中国（深圳）国际贸易单一窗口 - Google Chrome':
                # h 为想要放到最前面的窗口句柄
                win32gui.BringWindowToTop(h)
                shell = win32com.client.Dispatch("WScript.Shell")
                shell.SendKeys('%')

                # 如果被其他窗口遮挡，调用后放到最前面
                win32gui.SetForegroundWindow(h)

                # 解决被最小化的情况
                win32gui.ShowWindow(h, win32con.SW_MAXIMIZE)


def login_taxpage(browser):
    """
    程序开始执行，进入税费单界面等待rabbitmq返回待抓取信息
    """
    browser.get('https://sz.singlewindow.cn/dyck/')
    sleep(3)
    top_windows()
    sleep(10)


if __name__ == '__main__':
    create_debugwindow()