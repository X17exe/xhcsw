import json
import os
import random
import re
import sys
from time import sleep
import pika
import pyautogui
import requests
import win32api
import win32com.client
import win32con
import win32gui
import yagmail
import yaml
from selenium import webdriver
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver import ChromeOptions
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait


def read_yaml():
    """
    读取配置文件
    """
    with open('D:/workspace/xht_sw/tax_crawling/configure.yml', encoding="utf-8") as y:
    # with open('C:/dist/run/configure.yml', encoding="utf-8") as y:
        yaml_data = yaml.safe_load(y)
    return yaml_data


def upload_taxjson(detail_data):
    """
    上传税单json对象
    """
    # url = read_yaml()['upload_api']['tax_jsonapi']
    # header = {'Content-Type': 'application/json'}
    # json_data = detail_data
    # r = requests.post(url=url, headers=header, json=json_data)
    # code = int(r.json()['code'])
    # msg = r.json()['msg']
    # if code == 200:
    #     print("数据上传成功！")
    # else:
    #     print("数据上传失败,原因：%s" % (msg))
    #     upload_content = "税费数据上传失败！\n失败原因：%s,\n请求接口：%s,\n请求数据：%s" % (msg, url, detail_data)
    #     print(upload_content)
    print(detail_data)


def get_goods_number(line, browser):
    try:
        browser.find_element_by_xpath('//*[@id="taxationQueryNTable"]/tbody/tr[%s]/td[1]/label/input' % line).click()
    except:
        browser.find_element_by_xpath('//*[@id="taxationQuerySTable"]/tbody/tr[%s]/td[1]/label/input' % line).click()
    sleep(1)
    try:
        browser.find_element_by_xpath('//button[@id="taxationGoodsButton"]').click()  # 点击 税单货物信息
    except:
        browser.find_element_by_xpath('//button[@id="taxationGoodsSButton"]').click()  # 点击 税单货物信息
    WebDriverWait(browser, 100).until(lambda x: x.find_element_by_xpath('//*[@id="taxationGoods"]/div/div/div/div'
                                                                        '/div[1]/div[3]/div[1]').is_displayed())
    sleep(0.5)
    number_text = browser.find_element_by_xpath('//*[@id="taxationGoods"]/div/div/div/div/div[1]/div[3]/div[1]'
                                                '/span[@class="pagination-info"]').text
    data_number_list = re.findall(r"总共 (.*?) 条记录", number_text)
    data_number = int(data_number_list[0])

    return data_number


def set_details_one_page(browser):
    #  点击展开分页条数按钮
    browser.find_element_by_xpath('//*[@id="taxationGoods"]/div/div/div/div/div[1]/div[3]/div[1]'
                                  '/span[2]/span/button').click()
    sleep(1)
    #  点击选择最大的页面条数按钮
    browser.find_element_by_xpath('//*[@id="taxationGoods"]/div/div/div/div/div[1]/div[3]/div[1]'
                                  '/span[2]/span/ul/li[last()]').click()
    sleep(1)


def get_goods_detail_number(browser):
    #  获取数据总条数文本
    data_number_text = browser.find_element_by_xpath('//*[@id="taxationGoods"]/div/div/div/div/div[1]/div[3]'
                                                     '/div[1]/span[@class="pagination-info"]').text
    #  提取数据条数数量
    goods_data_number_list = re.findall(r"总共 (.*?) 条记录", data_number_text)
    goods_data_number = int(goods_data_number_list[0]) + 1

    return goods_data_number


def close_tax_windows_uncheck(line, browser):
    #  关闭税单货物信息窗口
    browser.find_element_by_xpath('//a[@class="layui-layer-ico layui-layer-close layui-layer-close1"]').click()
    #  取消勾选列表数据
    try:
        browser.find_element_by_xpath('//*[@id="taxationQueryNTable"]/tbody/tr[%s]/td[1]/label/input' % line).click()
    except:
        browser.find_element_by_xpath('//*[@id="taxationQuerySTable"]/tbody/tr[%s]/td[1]/label/input' % line).click()


def get_tax_category_name(line, browser):
    try:
        tax_name = browser.find_element_by_xpath('//*[@id="taxationQueryNTable"]/tbody/tr[%s]/td[6]' % line).text
    except:
        tax_name = browser.find_element_by_xpath('//*[@id="taxationQuerySTable"]/tbody/tr[%s]/td[6]' % line).text
    if '进口关税' in tax_name:
        tax_type = 'A'
    else:
        tax_type = 'L'

    return tax_type


def goods_detail_dict(tax_category, cus_declaration_no, line, browser):
    detail_dict = {}

    #  商品序号
    commodity_no = browser.find_element_by_xpath('//*[@id="taxationGoodsTable"]/tbody'
                                                 '/tr[%i]/td[3]' % line).text
    #  税号
    tax_no = browser.find_element_by_xpath('//*[@id="taxationGoodsTable"]/tbody'
                                           '/tr[%i]/td[4]' % line).text
    #  货名
    goods_name = browser.find_element_by_xpath('//*[@id="taxationGoodsTable"]/tbody'
                                               '/tr[%i]/td[5]' % line).text
    #  从价税率
    valorem_tax_rate = browser.find_element_by_xpath('//*[@id="taxationGoodsTable"]/tbody'
                                                     '/tr[%i]/td[6]' % line).text
    #  从量税率
    specific_tax_rate = browser.find_element_by_xpath('//*[@id="taxationGoodsTable"]/tbody'
                                                      '/tr[%i]/td[7]' % line).text
    #  税额
    tax_amount = browser.find_element_by_xpath('//*[@id="taxationGoodsTable"]/tbody'
                                               '/tr[%i]/td[8]' % line).text
    detail_dict["clTaxRate"] = float(specific_tax_rate)
    detail_dict["ctaxRate"] = float(valorem_tax_rate)
    detail_dict["customsCode"] = tax_no
    detail_dict["customsNo"] = cus_declaration_no
    detail_dict["itemNo"] = int(commodity_no)
    detail_dict["productCustomsName"] = goods_name
    detail_dict["taxMoney"] = float(tax_amount)
    detail_dict["taxType"] = tax_category

    return detail_dict


tax_no = "530120231010136424"


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


def a():
    top_windows()
    option = ChromeOptions()
    option.add_experimental_option("debuggerAddress", "127.0.0.1:9000")  # 打开已开启调试窗口
    option.add_argument('--no-sandbox')
    browser = webdriver.Chrome(options=option)
    browser.refresh()
    browser.implicitly_wait(5)
    browser.switch_to.frame(browser.find_element_by_xpath('//iframe[@name="layui-layer-iframe2"]'))
    sleep(1)
    browser.find_element_by_xpath('//img[@alt="关闭"]').click()
    browser.switch_to.default_content()  # 切回默认层
    browser.find_element_by_xpath('//span[text()="支付管理"]').click()
    sleep(1)
    browser.find_element_by_xpath('//a[text()="税费单支付"]').click()
    browser.switch_to.frame(browser.find_element_by_xpath('//iframe[@name="iframe2"]'))
    sleep(1)
    browser.find_element_by_xpath('//input[@id="entryIdN"]').clear()
    browser.find_element_by_xpath('//input[@id="entryIdN"]').send_keys(tax_no)
    browser.find_element_by_xpath('//button[@id="taxationQueryBtnN"]').click()  # 未支付页面查询按钮
    sleep(5)
    browser.find_element_by_xpath('//*[@id="tab-1"]/div/div/div[3]/div/div/div/div[1]/div[2]/div[1]'
                                  '/table/thead/tr/th[1]/div[1]/label/input').click()  # 勾选数据
    sleep(1)
    browser.find_element_by_xpath('//button[@id="taxationPrintNButton"]').click()  # 点击预览打印
    sleep(5)
    window = browser.window_handles
    window_number = len(window)
    sleep(1)
    if window_number == 2:
        pyautogui.hotkey('ctrl', 'w')  # 关闭下载页面
        sleep(1)
        browser.find_element_by_xpath('//*[@id="tab-1"]/div/div/div[3]/div/div/div/div[1]/div[2]/div[1]'
                                      '/table/thead/tr/th[1]/div[1]/label/input').click()  # 取消勾选数据
        tax_category_num_text = browser.find_element_by_xpath(
            '//*[@id="tab-1"]/div/div/div[3]/div/div/div/div[1]'
            '/div[3]/div[1]/span[1]').text  # 税种数量
        tax_category_num_list = re.findall(r"总共 (.*?) 条记录", tax_category_num_text)
        tax_category_num = int(tax_category_num_list[0])

        if tax_category_num == 2:  # 能进说明报关单有两条税种
            '''
            勾选列表第一条数据
            '''
            #  获取报关单号
            cus_no = browser.find_element_by_xpath(
                '//*[@id="taxationQueryNTable"]/tbody/tr[1]/td[3]').text
            first_tax_detail_list = []
            first_tax_type = get_tax_category_name(1, browser)
            one_data_number = get_goods_number(1, browser)
            if one_data_number > 10:
                # 说明不止十条数据，转为一页显示
                sleep(3)
                set_details_one_page(browser)
            first_data_number = get_goods_detail_number(browser)
            #  循环读取税费货物信息数据
            for detail_no in range(1, first_data_number):
                first_tax_goods_detail_dict = goods_detail_dict(first_tax_type, cus_no,
                                                                detail_no, browser)
                first_tax_detail_list.append(first_tax_goods_detail_dict)
            first_tax_detail_json = json.dumps(first_tax_detail_list, ensure_ascii=False)
            upload_taxjson(first_tax_detail_json)
            close_tax_windows_uncheck(1, browser)

            '''
            勾选列表第二条数据
            '''
            second_tax_detail_list = []
            second_tax_type = get_tax_category_name(2, browser)
            two_data_number = get_goods_number(2, browser)
            if two_data_number > 10:
                # 说明不止十条数据，转为一页显示
                sleep(3)
                set_details_one_page(browser)
            second_data_number = get_goods_detail_number(browser)
            #  循环读取税费货物信息数据
            for second_detail_no in range(1, second_data_number):
                second_tax_goods_detail_dict = goods_detail_dict(second_tax_type, cus_no,
                                                                 second_detail_no, browser)
                second_tax_detail_list.append(second_tax_goods_detail_dict)
            second_tax_detail_json = json.dumps(second_tax_detail_list, ensure_ascii=False)
            upload_taxjson(second_tax_detail_json)
            close_tax_windows_uncheck(2, browser)

        else:  # 能进说明报关单只有一个税种
            '''
            勾选列表第一条数据
            '''
            #  获取报关单号
            cus_no = browser.find_element_by_xpath(
                '//*[@id="taxationQueryNTable"]/tbody/tr[1]/td[3]').text
            first_tax_detail_list = []
            first_tax_type = get_tax_category_name(1, browser)
            one_data_number = get_goods_number(1, browser)
            if one_data_number > 10:
                # 说明不止十条数据，转为一页显示
                sleep(3)
                set_details_one_page(browser)
            first_data_number = get_goods_detail_number(browser)
            #  循环读取税费货物信息数据
            for detail_no in range(1, first_data_number):
                first_tax_goods_detail_dict = goods_detail_dict(first_tax_type, cus_no,
                                                                detail_no, browser)
                first_tax_detail_list.append(first_tax_goods_detail_dict)
            first_tax_detail_json = json.dumps(first_tax_detail_list, ensure_ascii=False)
            upload_taxjson(first_tax_detail_json)
            close_tax_windows_uncheck(1, browser)
    #  去已支付界面查询税单
    else:
        # browser.switch_to.frame(browser.find_element_by_xpath('//iframe[@name="iframe2"]'))
        browser.find_element_by_xpath('//a[@id="payFinishTab"]').click()
        sleep(1)
        browser.find_element_by_xpath('//input[@id="entryIdS"]').clear()
        browser.find_element_by_xpath('//input[@id="entryIdS"]').send_keys(tax_no)
        browser.find_element_by_xpath('//button[@id="taxationQueryBtnS"]').click()  # 未支付页面查询按钮
        sleep(5)
        browser.find_element_by_xpath('//*[@id="tab-3"]/div/div/div[3]/div/div/div/div[1]/div[2]'
                                      '/div[1]/table/thead/tr/th[1]/div[1]/label/input').click()  # 勾选数据
        sleep(1)
        browser.find_element_by_xpath('//button[@id="taxationPrintSButton"]').click()  # 点击预览打印
        sleep(5)
        window = browser.window_handles
        window_number = len(window)
        if window_number == 2:
            pyautogui.hotkey('ctrl', 'w')  # 关闭下载页面
            sleep(1)
            browser.find_element_by_xpath('//*[@id="tab-3"]/div/div/div[3]/div/div/div/div[1]/div[2]'
                                          '/div[1]/table/thead/tr/th[1]/div[1]/label/input').click()  # 取消勾选数据
            tax_category_num_text = browser.find_element_by_xpath(
                '//*[@id="tab-3"]/div/div/div[3]/div/div/div/div[1]/div[3]/div[1]/span[1]').text  # 税种数量
            tax_category_num_list = re.findall(r"总共 (.*?) 条记录", tax_category_num_text)
            tax_category_num = int(tax_category_num_list[0])

            if tax_category_num == 2:  # 能进说明报关单有两条税种
                '''
                勾选列表第一条数据
                '''
                #  获取报关单号
                cus_no = browser.find_element_by_xpath(
                    '//*[@id="taxationQuerySTable"]/tbody/tr[1]/td[3]').text
                first_tax_detail_list = []
                first_tax_type = get_tax_category_name(1, browser)
                one_data_number = get_goods_number(1, browser)
                if one_data_number > 10:
                    # 说明不止十条数据，转为一页显示
                    sleep(3)
                    set_details_one_page(browser)
                first_data_number = get_goods_detail_number(browser)
                #  循环读取税费货物信息数据
                for detail_no in range(1, first_data_number):
                    first_tax_goods_detail_dict = goods_detail_dict(first_tax_type, cus_no,
                                                                    detail_no, browser)
                    first_tax_detail_list.append(first_tax_goods_detail_dict)
                first_tax_detail_json = json.dumps(first_tax_detail_list, ensure_ascii=False)
                upload_taxjson(first_tax_detail_json)
                close_tax_windows_uncheck(1, browser)

                '''
                勾选列表第二条数据
                '''
                second_tax_detail_list = []
                second_tax_type = get_tax_category_name(2, browser)
                two_data_number = get_goods_number(2, browser)
                if two_data_number > 10:
                    # 说明不止十条数据，转为一页显示
                    sleep(3)
                    set_details_one_page(browser)
                second_data_number = get_goods_detail_number(browser)
                #  循环读取税费货物信息数据
                for second_detail_no in range(1, second_data_number):
                    second_tax_goods_detail_dict = goods_detail_dict(second_tax_type, cus_no,
                                                                     second_detail_no, browser)
                    second_tax_detail_list.append(second_tax_goods_detail_dict)
                second_tax_detail_json = json.dumps(second_tax_detail_list, ensure_ascii=False)
                upload_taxjson(second_tax_detail_json)
                close_tax_windows_uncheck(2, browser)

            else:  # 能进说明报关单只有一个税种
                '''
                勾选列表第一条数据
                '''
                #  获取报关单号
                cus_no = browser.find_element_by_xpath(
                    '//*[@id="taxationQuerySTable"]/tbody/tr[1]/td[3]').text
                first_tax_detail_list = []
                first_tax_type = get_tax_category_name(1, browser)
                one_data_number = get_goods_number(1, browser)
                if one_data_number > 10:
                    # 说明不止十条数据，转为一页显示
                    sleep(3)
                    set_details_one_page(browser)
                first_data_number = get_goods_detail_number(browser)
                #  循环读取税费货物信息数据
                for detail_no in range(1, first_data_number):
                    first_tax_goods_detail_dict = goods_detail_dict(first_tax_type, cus_no,
                                                                    detail_no, browser)
                    first_tax_detail_list.append(first_tax_goods_detail_dict)
                first_tax_detail_json = json.dumps(first_tax_detail_list, ensure_ascii=False)
                upload_taxjson(first_tax_detail_json)
                close_tax_windows_uncheck(1, browser)
            browser.find_element_by_xpath('//a[@id="unPayTab"]').click()
        else:
            browser.find_element_by_xpath('//a[@id="unPayTab"]').click()
            print("已支付税费未查得")


a()


# def except_send_email(content, ec):
#     """
#     发送程序运行异常邮件提示
#     """
#     yag = yagmail.SMTP(user='764178925@qq.com',
#                        password='rilvnbgrpaqabcce',
#                        host='smtp.qq.com',
#                        port=465)
#     contents = '%s\n%s' % (content, ec)  # 邮件内容
#     subject = 'grab_tax爬虫程序运行异常通知！'  # 邮件主题
#     receiver = read_yaml()['email']['receiver']  # 接收方邮箱账号
#     yag.send(receiver, subject, contents)
#     yag.close()


# 连接rabbit
# credentials = pika.PlainCredentials(read_yaml()['rabbitmq']['user'], read_yaml()['rabbitmq']['password'])
# parameters = pika.ConnectionParameters(read_yaml()['rabbitmq']['host'],
#                                        read_yaml()['rabbitmq']['port'],
#                                        '/',
#                                        credentials)
# connection = pika.BlockingConnection(parameters)
# channel = connection.channel()
# # 声明direct
# channel.exchange_declare(exchange=read_yaml()['rabbitmq']['exchange'], exchange_type='direct', durable=True)
# channel.queue_declare(queue=read_yaml()['rabbitmq']['queue'], durable=True)
# # 队列绑定交换机
# channel.queue_bind(exchange=read_yaml()['rabbitmq']['exchange'], queue=read_yaml()['rabbitmq']['queue'],
#                    routing_key=read_yaml()['rabbitmq']['routing_key'])
#
#
# def callback(ch, method, properties, body):
#     tax_no = body.decode()
#     print(tax_no)
#     ch.basic_ack(delivery_tag=method.delivery_tag)
#
#
# channel.basic_consume(queue=read_yaml()['rabbitmq']['queue'],
#                       auto_ack=False,
#                       on_message_callback=callback)
#
# channel.start_consuming()


# url = 'https://api.delchannel.com/customs/hgBill/declare/uploadTaxJson'
# header = {'Content-Type': 'application/json;charset=UTF-8'}
# r = requests.post(url=url, headers=header, data=detail_data.encode("utf-8"))
# code = int(r.json()['code'])
# msg = r.json()['msg']
# if code == 200:
#     print("数据上传成功！")
# else:
#     print("数据上传失败,原因：%s" % (msg))


# goodsfile_save_path = read_yaml()['localfile']['save_path']
# uploaded_path = read_yaml()['localfile']['uploaded_path']
# uploadfail_path = read_yaml()['localfile']['uploadfail_path']
# filelist = os.listdir(goodsfile_save_path)
# filelist.sort(key=lambda fn: os.path.getmtime(goodsfile_save_path + '\\' + fn))
# name = ''.join(filelist[-1])
# goods_file_name = goodsfile_save_path + name
#  接口参数
# filetype = 1
# tax_no = 530120231010064150
# url = 'https://api.delchannel.com/customs/hgBill/declare/uploadBillMultipart/' + str(tax_no) + "/%s" % filetype
# content = open("C:\\Users\\admin\\Desktop\\I20230001011631967.pdf", 'rb')
# files = {'file': ("I20230001011631967.pdf", content, 'pdf')}
# data = {'Content-Disposition': 'form-data', 'Content-Type': 'application/pdf'}
# r = requests.post(url=url, data=data, files=files)
# content.close()
# code = int(r.json()['code'])
# msg = r.json()['msg']
# error_text = r.text
# print(code, msg)
