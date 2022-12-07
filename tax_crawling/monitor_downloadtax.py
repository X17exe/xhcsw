from selenium import webdriver
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver import ChromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
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
import re
import json
import win32gui
import win32con
import win32com.client


def read_yaml():
    """
    读取配置文件
    """
    try:
        # with open('D:/workspace/xht_sw/tax_crawling/configure.yml', encoding="utf-8") as y:
        with open('C:/dist/run/configure.yml', encoding="utf-8") as y:
            yaml_data = yaml.safe_load(y)
        return yaml_data
    except BaseException as t:
        config_content = "配置文件读取错误，请将配置文件放置在'C:/dist/run/'目录下"
        except_send_email(content=config_content, ec=t)


def except_send_email(content, ec):
    """
    发送程序运行异常邮件提示
    """
    yag = yagmail.SMTP(user='764178925@qq.com',
                       password='rilvnbgrpaqabcce',
                       host='smtp.qq.com',
                       port=465)
    contents = '%s\n%s' % (content, ec)  # 邮件内容
    subject = 'grab_tax爬虫程序运行异常通知！'  # 邮件主题
    receiver = read_yaml()['email']['receiver']  # 接收方邮箱账号
    yag.send(receiver, subject, contents)
    yag.close()


def create_files():
    #  创建脚本抓取文件的本地保存路径
    file1 = os.path.exists(read_yaml()['localfile']['save_path'])
    if file1:
        pass
    else:
        os.makedirs(read_yaml()['localfile']['save_path'])

    file2 = os.path.exists(read_yaml()['localfile']['uploaded_path'])
    if file2:
        pass
    else:
        os.makedirs(read_yaml()['localfile']['uploaded_path'])

    file3 = os.path.exists(read_yaml()['localfile']['uploadfail_path'])
    if file3:
        pass
    else:
        os.makedirs(read_yaml()['localfile']['uploadfail_path'])

    file4 = os.path.exists('C:/chromedebug/')
    if file4:
        pass
    else:
        os.makedirs('C:/chromedebug/')


def create_debugwindow():
    #  创建chrome浏览器调试窗口
    cmd = read_yaml()['localfile']['chromedebugwindow']
    win32api.ShellExecute(0, 'open', cmd, '', '', 1)
    sleep(5)
    pyautogui.hotkey('alt', ' ', 'x')  # 最大化浏览器窗口
    sleep(2)
    pyautogui.press('esc')

    option = ChromeOptions()
    option.add_experimental_option("debuggerAddress", "127.0.0.1:9000")  # 打开已开启调试窗口
    option.add_argument('--no-sandbox')
    browser = webdriver.Chrome(options=option)
    win = browser.window_handles
    win_number = len(win)
    if win_number == 1:
        login_taxpage(browser)
    else:
        pass


def close_inform(browser):
    """
    关闭弹出页面
    :param browser: webdriver
    """
    try:
        close = browser.find_element_by_xpath('//span[@class="layui-layer-setwin"]/a')
        close.click()
    except BaseException:
        pass


handle_list = []
relogin_handle_list = []


def login_taxpage(browser):
    """
    程序开始执行，进入税费单界面等待rabbitmq返回待抓取信息
    """
    try:
        handle_list.clear()
        relogin_handle_list.clear()
        browser.get('https://sz.singlewindow.cn/dyck/')
        browser.execute_script('() =>{ Object.defineProperties(navigator,{ webdriver:{ get: () => false } }) }')
        close_inform(browser)  # 关闭弹出通知
        singe_index = browser.current_window_handle
        handle_list.append(singe_index)
        relogin_handle_list.append(singe_index)

        browser.switch_to.frame('loginIframe1')  # 切入frame层

        # 判断页面是否加载完成
        WebDriverWait(browser, 100).until(lambda x: x.find_element_by_xpath('//a[@id="ic_card"]').is_displayed())
        login_button = browser.find_element_by_xpath('//a[@id="ic_card"]')
        login_button.click()
        browser.switch_to.default_content()  # 切回默认层

        browser.switch_to.frame('loginIframe2')
        WebDriverWait(browser, 100).until(lambda x: x.find_element_by_xpath('//input[@id="password"]').is_displayed())
        password_input = browser.find_element_by_xpath('//input[@id="password"]')
        password_input.send_keys(read_yaml()['single_widow']['password'])
        login = browser.find_element_by_id('loginbutton')
        login.click()
        sleep(10)
        browser.switch_to.default_content()
        sleep(1)
        close_inform(browser)  # 关闭弹出通知

        #  进入税单界面
        browser.switch_to.frame(browser.find_element_by_xpath('//div[@id="content"]/iframe'))
        taxes_transact = browser.find_element_by_xpath('//li[contains(text(), "税费办理")]')
        taxes_transact.click()
        sleep(5)
        browser.switch_to.window(browser.window_handles[1])
        sleep(1)
        tax_index = browser.current_window_handle
        handle_list.append(tax_index)
        relogin_handle_list.append(tax_index)
        tax_should_url = 'https://sz.singlewindow.cn/dyck/swProxy/deskserver/sw/deskIndex?menu_id=spl'
        tax_current_url = browser.current_url
        if tax_current_url == tax_should_url:
            pass
        else:
            tax_ex_content = "监测到程序运行异常，请检查操作员卡是否插好，并按照使用手册重新运行程序！"
            except_send_email(content=tax_ex_content, ec=None)
            sys.exit(0)

        #  进入货物申报界面
        browser.switch_to.window(browser.window_handles[0])  # 返回单一首页
        sleep(0.5)
        browser.switch_to.frame(browser.find_element_by_xpath('//div[@id="content"]/iframe'))
        taxes_transact = browser.find_element_by_xpath('//li[contains(text(), "货物申报")]')
        taxes_transact.click()
        browser.find_element_by_xpath(
            '//a[@data-href="http://sz.singlewindow.cn/dyck/swProxy/deskserver/sw/deskIndex?menu_id=dec001"]').click()
        sleep(5)
        browser.switch_to.window(browser.window_handles[-1])
        sleep(1)
        goods_index = browser.current_window_handle
        handle_list.append(goods_index)
        relogin_handle_list.append(goods_index)
        goods_should_url = 'https://sz.singlewindow.cn/dyck/swProxy/deskserver/sw/deskIndex?menu_id=dec001'
        goods_current_url = browser.current_url
        if goods_current_url == goods_should_url:
            pass
        else:
            goods_ex_content = "监测到程序运行异常，请检查操作员卡是否插好，并按照使用手册重新运行程序！"
            except_send_email(content=goods_ex_content, ec=None)
            sys.exit(0)

        browser.service.stop()
    except BaseException as o:
        login_content = "单一窗口未登录成功，请阅读错误消息，关闭当前程序然后重新运行程序"
        except_send_email(content=login_content, ec=o)


# 连接rabbit
credentials = pika.PlainCredentials(read_yaml()['rabbitmq']['user'], read_yaml()['rabbitmq']['password'])

parameters = pika.ConnectionParameters(read_yaml()['rabbitmq']['host'],
                                       read_yaml()['rabbitmq']['port'],
                                       '/',
                                       credentials)

connection = pika.BlockingConnection(parameters)

# connection = pika.BlockingConnection(pika.ConnectionParameters(host=read_yaml()['rabbitmq']['host'],
#                                                                port=read_yaml()['rabbitmq']['port']))
channel = connection.channel()

# 声明direct
channel.exchange_declare(exchange=read_yaml()['rabbitmq']['exchange'], exchange_type='direct', durable=True)

channel.queue_declare(queue=read_yaml()['rabbitmq']['queue'], durable=True)

# 队列绑定交换机
channel.queue_bind(exchange=read_yaml()['rabbitmq']['exchange'], queue=read_yaml()['rabbitmq']['queue'],
                   routing_key=read_yaml()['rabbitmq']['routing_key'])


def rand_sleep():
    """
    随机生成等待时间，等待该时间后再去抓取税单
    """
    time = random.randint(1, 5)
    # time = random.randint(30, 60)
    return time


def upload_taxjson(detail_data):
    """
    上传税单json对象
    """
    url = read_yaml()['upload_api']['tax_json_api']
    header = {'Content-Type': 'application/json'}
    json_data = detail_data
    r = requests.post(url=url, headers=header, json=json_data)
    code = int(r.json()['code'])
    msg = r.json()['msg']
    if code == 200:
        print("数据上传成功！")
    else:
        print("数据上传失败,原因：%s" % (msg))
        upload_content = "税费数据上传失败！\n失败原因：%s,\n请求接口：%s,\n请求数据：%s" % (msg, url, detail_data)
        except_send_email(content=upload_content, ec=None)


def upload_tax_goods_file(tax_no, filetype):
    """
    上传税单、货物单二进制文件流
    """
    #  文件参数
    goodsfile_save_path = read_yaml()['localfile']['save_path']
    uploaded_path = read_yaml()['localfile']['uploaded_path']
    uploadfail_path = read_yaml()['localfile']['uploadfail_path']
    filelist = os.listdir(goodsfile_save_path)
    filelist.sort(key=lambda fn: os.path.getmtime(goodsfile_save_path + '\\' + fn))
    name = ''.join(filelist[-1])
    goods_file_name = goodsfile_save_path + name
    #  接口参数
    url = read_yaml()['upload_api']['tax_goods_api'] + tax_no + "/%s" % filetype
    content = open(goods_file_name, 'rb')
    files = {'file': (name, content, 'pdf')}
    data = {'Content-Disposition': 'form-data', 'Content-Type': 'application/pdf'}
    r = requests.post(url=url, data=data, files=files)
    content.close()
    code = int(r.json()['code'])
    msg = r.json()['msg']
    if code == 200:
        shutil.move(goods_file_name, uploaded_path)
        print("文件 %s 上传成功！" % name)
    else:
        print("文件 %s 上传失败,原因：%s" % (name, msg))
        uploadfail_content = "税费文件 %s 上传失败，请在 %s 文件中查看!\n失败原因：%s" % (name, uploadfail_path, msg)
        except_send_email(content=uploadfail_content, ec=None)
        shutil.move(goods_file_name, uploadfail_path)


def get_goods_number(line, browser):
    # 勾选数据
    browser.find_element_by_xpath('//*[@id="taxationQueryNTable"]/tbody/tr[%s]/td[1]/label/input' % line).click()
    sleep(1)
    browser.find_element_by_xpath('//button[@id="taxationGoodsButton"]').click()  # 点击税单货物信息
    #  校验税单货物信息列表是否渲染完成
    WebDriverWait(browser, 100).until(lambda x: x.find_element_by_xpath('//*[@id="taxationGoods"]/div/div/div/div'
                                                                        '/div[1]/div[3]/div[1]').is_displayed())
    sleep(1)
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
    sleep(2)


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
    sleep(1)
    #  取消勾选列表数据
    browser.find_element_by_xpath('//*[@id="taxationQueryNTable"]/tbody/tr[%s]/td[1]/label/input' % line).click()


def get_tax_category_name(line, browser):
    #  获取税种
    tax_name = browser.find_element_by_xpath('//*[@id="taxationQueryNTable"]/tbody/tr[%s]/td[6]' % line).text
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


def go_to_download():
    """
    税单下载页面，操作下载
    """
    sleep(2)
    pyautogui.hotkey('ctrl', 's')  # 保存
    sleep(1)
    pyautogui.press('left')
    sleep(1)
    pyautogui.write(read_yaml()['localfile']['file_path_prefix'])
    sleep(2)
    pyautogui.press('shift')  # 防止中文输入法
    sleep(2)
    pyautogui.press('enter')  # 点击保存本地
    sleep(1)
    pyautogui.hotkey('ctrl', 'w')  # 关闭下载页面
    sleep(1)


def top_windows():
    """
    置顶chrome_debug浏览器窗口
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


# 回调爬取
def callback(ch, method, properties, body):
    """
    根据回调的报关单号，执行抓取任务
    """
    try:
        top_windows()
        tax_no = body.decode()
        option1 = ChromeOptions()
        option1.add_experimental_option("debuggerAddress", "127.0.0.1:9000")  # 打开已开启调试窗口
        option1.add_argument('--no-sandbox')
        browser = webdriver.Chrome(options=option1)
        sleep(1)
        refresh_code = "123456789"
        if tax_no != refresh_code:
            grab_mode = read_yaml()['grab_mode']['mode']
            if int(grab_mode) == 1:
                #  货物单抓取
                browser.switch_to.window(handle_list[2])  # 切换到货物申报界面
                sleep(rand_sleep())
                browser.refresh()
                browser.implicitly_wait(5)
                # 判断cookie是否失效，失效则重新登录
                goods_url = 'https://sz.singlewindow.cn/dyck/swProxy/deskserver/sw/deskIndex?menu_id=dec001'
                goods_nowurl = browser.current_url
                if goods_nowurl != goods_url:
                    pyautogui.hotkey('alt', 'F4')
                    browser.service.stop()
                    create_debugwindow()
                    sleep(2)
                    message = "报关单 %s 因程序重启导致抓取失败，请将该报关单号反馈给技术人员重新抓取" % tax_no
                    except_send_email(message, None)
                elif goods_nowurl == goods_url:
                    browser.find_element_by_xpath('//span[text()="综合查询"]').click()
                    sleep(1)
                    browser.find_element_by_xpath('//a[text()= "报关数据查询"]').click()
                    WebDriverWait(browser, 10, 0.5).until(expected_conditions.presence_of_element_located(
                        (By.XPATH, '//iframe[@name="iframe01"]')))
                    sleep(2)
                    # 切入列表页面iframe
                    browser.switch_to.frame(browser.find_element_by_xpath('//iframe[@name="iframe01"]'))
                    browser.find_element_by_xpath('//input[@id="entryId1"]').send_keys(tax_no)  # 输入待抓取单号
                    sleep(1)
                    browser.find_element_by_xpath('//input[@id="operateDate2"]').click()  # 选择本周
                    sleep(1)
                    browser.find_element_by_xpath('//button[@id="decQuery"]').click()  # 点击查询
                    WebDriverWait(browser, 100).until(
                        lambda x: x.find_element_by_xpath(
                            '//*[@id="cus_declare_table_div"]/div[1]/div[2]/div[4]/div[1]').is_displayed())
                    sleep(1)
                    browser.find_element_by_xpath('//button[@id="decPdfPrint"]').click()  # 打印
                    WebDriverWait(browser, 100).until(
                        lambda x: x.find_element_by_xpath('//a[text()= "打印预览"]').is_displayed())
                    sleep(1)
                    browser.find_element_by_xpath('//input[@id="printSort3"]').click()  # 勾选商品附加页
                    sleep(1)
                    browser.find_element_by_xpath('//a[text()= "打印预览"]').click()
                    sleep(5)
                    go_to_download()
                    sleep(10)
                    upload_tax_goods_file(str(tax_no), 1)
                    sleep(3)

                    #  税单抓取
                    browser.switch_to.window(handle_list[1])  # 切换到税费window
                    browser.refresh()
                    browser.implicitly_wait(5)
                    # 切入iframe
                    browser.switch_to.frame(browser.find_element_by_xpath('//iframe[@name="layui-layer-iframe2"]'))
                    sleep(1)
                    browser.find_element_by_xpath('//img[@alt="关闭"]').click()
                    browser.switch_to.default_content()  # 切回默认层
                    browser.find_element_by_xpath('//span[text()="支付管理"]').click()
                    sleep(1)
                    browser.find_element_by_xpath('//a[text()="税费单支付"]').click()
                    WebDriverWait(browser, 10, 0.5).until(expected_conditions.presence_of_element_located(
                        (By.XPATH, '//*[@id="content-main"]/iframe[2]')))
                    # 切入列表页面iframe
                    browser.switch_to.frame(browser.find_element_by_xpath('//iframe[@name="iframe2"]'))
                    sleep(2)
                    browser.find_element_by_xpath('//input[@id="entryIdN"]').clear()
                    browser.find_element_by_xpath('//input[@id="entryIdN"]').send_keys(tax_no)
                    browser.find_element_by_xpath('//button[@id="taxationQueryBtnN"]').click()  # 未支付页面查询按钮
                    WebDriverWait(browser, 100).until(lambda x: x.find_element_by_xpath(
                        '//*[@id="tab-1"]/div/div/div[3]/div/div/div/div[1]/div[3]/div[1]').is_displayed())
                    sleep(2)
                    browser.find_element_by_xpath('(//input[@name="btSelectAll"])[1]').click()  # 勾选数据
                    browser.find_element_by_xpath('//button[@id="taxationPrintNButton"]').click()  # 点击预览打印
                    sleep(5)
                    window = browser.window_handles
                    window_number = len(window)
                    if window_number == 4:
                        go_to_download()
                        sleep(10)
                        upload_tax_goods_file(str(tax_no), 2)
                        sleep(1)
                        browser.find_element_by_xpath('(//input[@name="btSelectAll"])[1]').click()  # 取消勾选数据
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
                                set_details_one_page(browser)
                            first_data_number = get_goods_detail_number(browser)
                            #  循环读取税费货物信息数据
                            for detail_no in range(0, first_data_number):
                                if detail_no == 0:
                                    continue
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
                                set_details_one_page(browser)
                            second_data_number = get_goods_detail_number(browser)
                            #  循环读取税费货物信息数据
                            for second_detail_no in range(0, second_data_number):
                                if second_detail_no == 0:
                                    continue
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
                                set_details_one_page(browser)
                            first_data_number = get_goods_detail_number(browser)
                            #  循环读取税费货物信息数据
                            for detail_no in range(0, first_data_number):
                                if detail_no == 0:
                                    continue
                                first_tax_goods_detail_dict = goods_detail_dict(first_tax_type, cus_no,
                                                                                detail_no, browser)
                                first_tax_detail_list.append(first_tax_goods_detail_dict)
                            first_tax_detail_json = json.dumps(first_tax_detail_list, ensure_ascii=False)
                            upload_taxjson(first_tax_detail_json)
                            close_tax_windows_uncheck(1, browser)
                        ch.basic_ack(delivery_tag=method.delivery_tag)
                    else:
                        browser.refresh()
                        except_content = '报关单 %i 未在单一系统未支付税费模块查询到相关税费文件，请尽快核实' % tax_no
                        except_send_email(except_content, None)
                        ch.basic_ack(delivery_tag=method.delivery_tag)

            elif int(grab_mode) == 2:
                # 税单抓取
                browser.switch_to.window(handle_list[1])  # 切换到税费window
                sleep(rand_sleep())
                browser.refresh()
                browser.implicitly_wait(5)
                # 判断cookie是否失效，失效则重新登录
                tax_url = 'https://sz.singlewindow.cn/dyck/swProxy/deskserver/sw/deskIndex?menu_id=spl'
                tax_nowurl = browser.current_url
                if tax_nowurl != tax_url:
                    pyautogui.hotkey('alt', 'F4')
                    browser.service.stop()
                    create_debugwindow()
                    sleep(2)
                    message = "报关单 %s 因程序重启导致抓取失败，请将该报关单号反馈给技术人员重新抓取" % tax_no
                    except_send_email(message, None)
                elif tax_nowurl == tax_url:
                    browser.switch_to.frame(
                        browser.find_element_by_xpath('//iframe[@name="layui-layer-iframe2"]'))  # 切入iframe
                    sleep(1)
                    browser.find_element_by_xpath('//img[@alt="关闭"]').click()
                    browser.switch_to.default_content()  # 切回默认层
                    browser.find_element_by_xpath('//span[text()="支付管理"]').click()
                    sleep(1)
                    browser.find_element_by_xpath('//a[text()="税费单支付"]').click()
                    WebDriverWait(browser, 10, 0.5).until(expected_conditions.presence_of_element_located(
                        (By.XPATH, '//*[@id="content-main"]/iframe[2]')))
                    # 切入列表页面iframe
                    browser.switch_to.frame(browser.find_element_by_xpath('//iframe[@name="iframe2"]'))
                    sleep(2)
                    browser.find_element_by_xpath('//input[@id="entryIdN"]').clear()
                    browser.find_element_by_xpath('//input[@id="entryIdN"]').send_keys(tax_no)
                    browser.find_element_by_xpath('//button[@id="taxationQueryBtnN"]').click()  # 未支付页面查询按钮
                    WebDriverWait(browser, 100).until(lambda x: x.find_element_by_xpath(
                        '//*[@id="tab-1"]/div/div/div[3]/div/div/div/div[1]/div[3]/div[1]').is_displayed())
                    sleep(2)
                    browser.find_element_by_xpath('(//input[@name="btSelectAll"])[1]').click()  # 勾选数据
                    browser.find_element_by_xpath('//button[@id="taxationPrintNButton"]').click()  # 点击预览打印
                    sleep(5)
                    window = browser.window_handles
                    window_number = len(window)
                    if window_number == 4:
                        go_to_download()
                        sleep(10)
                        upload_tax_goods_file(str(tax_no), 2)
                        sleep(1)
                        browser.find_element_by_xpath('(//input[@name="btSelectAll"])[1]').click()  # 取消勾选数据
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
                                set_details_one_page(browser)
                            first_data_number = get_goods_detail_number(browser)
                            #  循环读取税费货物信息数据
                            for detail_no in range(0, first_data_number):
                                if detail_no == 0:
                                    continue
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
                                set_details_one_page(browser)
                            second_data_number = get_goods_detail_number(browser)
                            #  循环读取税费货物信息数据
                            for second_detail_no in range(0, second_data_number):
                                if second_detail_no == 0:
                                    continue
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
                                set_details_one_page(browser)
                            first_data_number = get_goods_detail_number(browser)
                            #  循环读取税费货物信息数据
                            for detail_no in range(0, first_data_number):
                                if detail_no == 0:
                                    continue
                                first_tax_goods_detail_dict = goods_detail_dict(first_tax_type, cus_no,
                                                                                detail_no, browser)
                                first_tax_detail_list.append(first_tax_goods_detail_dict)
                            first_tax_detail_json = json.dumps(first_tax_detail_list, ensure_ascii=False)
                            upload_taxjson(first_tax_detail_json)
                            close_tax_windows_uncheck(1, browser)
                        ch.basic_ack(delivery_tag=method.delivery_tag)
                    else:
                        browser.refresh()
                        except_content = '报关单 %i 未在单一系统未支付税费模块查询到相关税费文件，请尽快核实' % tax_no
                        except_send_email(except_content, None)
                        ch.basic_ack(delivery_tag=method.delivery_tag)

            elif int(grab_mode) == 3:
                #  货物单抓取
                browser.switch_to.window(handle_list[2])  # 切换到货物申报界面
                sleep(rand_sleep())
                browser.refresh()
                browser.implicitly_wait(5)
                # 判断cookie是否失效，失效则重新登录
                goods_url = 'https://sz.singlewindow.cn/dyck/swProxy/deskserver/sw/deskIndex?menu_id=dec001'
                goods_nowurl = browser.current_url
                if goods_nowurl != goods_url:
                    pyautogui.hotkey('alt', 'F4')
                    browser.service.stop()
                    create_debugwindow()
                    sleep(2)
                    message = "报关单 %s 因程序重启导致抓取失败，请将该报关单号反馈给技术人员重新抓取" % tax_no
                    except_send_email(message, None)
                elif goods_nowurl == goods_url:
                    browser.find_element_by_xpath('//span[text()="综合查询"]').click()
                    sleep(1)
                    browser.find_element_by_xpath('//a[text()= "报关数据查询"]').click()
                    WebDriverWait(browser, 10, 0.5).until(expected_conditions.presence_of_element_located(
                        (By.XPATH, '//iframe[@name="iframe01"]')))
                    # 切入列表页面iframe
                    browser.switch_to.frame(browser.find_element_by_xpath('//iframe[@name="iframe01"]'))
                    browser.find_element_by_xpath('//input[@id="entryId1"]').send_keys(tax_no)  # 输入待抓取单号
                    sleep(1)
                    browser.find_element_by_xpath('//input[@id="operateDate2"]').click()  # 选择本周
                    sleep(1)
                    browser.find_element_by_xpath('//button[@id="decQuery"]').click()  # 点击查询
                    WebDriverWait(browser, 100).until(
                        lambda x: x.find_element_by_xpath(
                            '//*[@id="cus_declare_table_div"]/div[1]/div[2]/div[4]/div[1]').is_displayed())
                    sleep(1)
                    browser.find_element_by_xpath('//button[@id="decPdfPrint"]').click()  # 打印
                    WebDriverWait(browser, 100).until(
                        lambda x: x.find_element_by_xpath('//a[text()= "打印预览"]').is_displayed())
                    sleep(1)
                    browser.find_element_by_xpath('//input[@id="printSort3"]').click()  # 勾选商品附加页
                    sleep(1)
                    browser.find_element_by_xpath('//a[text()= "打印预览"]').click()
                    sleep(5)
                    go_to_download()
                    sleep(10)
                    upload_tax_goods_file(str(tax_no), 1)
                    sleep(3)
                    ch.basic_ack(delivery_tag=method.delivery_tag)

        else:
            browser.switch_to.window(handle_list[0])
            browser.refresh()
            browser.implicitly_wait(5)
            browser.switch_to.window(handle_list[1])
            browser.refresh()
            browser.implicitly_wait(5)
            browser.switch_to.window(handle_list[2])
            browser.refresh()
            browser.implicitly_wait(5)
            gods_url = 'https://sz.singlewindow.cn/dyck/swProxy/deskserver/sw/deskIndex?menu_id=dec001'
            gods_now_url = browser.current_url
            if gods_now_url != gods_url:
                pyautogui.hotkey('alt', 'F4')
                browser.service.stop()
                create_debugwindow()
                sleep(3)
            ch.basic_ack(delivery_tag=method.delivery_tag)

        browser.service.stop()
    except BaseException as r:
        exception_content = "抓取程序异常！"
        except_send_email(exception_content, r)


# 监听队列参数
channel.basic_consume(queue=read_yaml()['rabbitmq']['queue'],
                      auto_ack=False,
                      on_message_callback=callback)
print('爬虫程序运行中，如果想退出，请直接关闭浏览器和程序窗口')

