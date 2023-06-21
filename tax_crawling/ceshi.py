from selenium import webdriver
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver import ChromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.webdriver.common.action_chains import ActionChains
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
        with open('D:/workspace/xht_sw/tax_crawling/configure.yml', encoding="utf-8") as y:
        # with open('C:/dist/run/configure.yml', encoding="utf-8") as y:
            yaml_data = yaml.safe_load(y)
        return yaml_data
    except BaseException as t:
        config_content = "配置文件读取错误，请将配置文件放置在'C:/dist/run/'目录下"
        except_send_email(content=config_content, ec=t)
        print(t)


def except_send_email(content, ec):
    try:
        """
        发送程序运行异常邮件提示
        """
        yag = yagmail.SMTP(user='764178925@qq.com',
                           password='rilvnbgrpaqabcce',
                           host='smtp.qq.com',
                           port=465)
        contents = '%s\n%s' % (content, ec)  # 邮件内容
        subject = 'grab_tax爬虫程序运行提示！'  # 邮件主题
        receiver = read_yaml()['email']['receiver']  # 接收方邮箱账号
        yag.send(receiver, subject, contents)
        yag.close()
    except BaseException as o:
        print(o)


def create_files():
    #  创建脚本抓取文件的本地保存路径
    file1 = os.path.exists(read_yaml()['localfile']['save_path'])
    if file1:
        pass
    else:
        os.makedirs(read_yaml()['localfile']['save_path'])

    # file2 = os.path.exists(read_yaml()['localfile']['uploaded_path'])
    # if file2:
    #     pass
    # else:
    #     os.makedirs(read_yaml()['localfile']['uploaded_path'])
    #
    # file3 = os.path.exists(read_yaml()['localfile']['uploadfail_path'])
    # if file3:
    #     pass
    # else:
    #     os.makedirs(read_yaml()['localfile']['uploadfail_path'])

    file4 = os.path.exists('C:/chromedebug/')
    if file4:
        pass
    else:
        os.makedirs('C:/chromedebug/')


def create_debugwindow():
    #  创建chrome浏览器调试窗口,创建窗口前先杀掉可能存在的debug窗口
    # os.system('powershell -command "Get-Process chrome | ForEach-Object { $_.CloseMainWindow() | Out-Null}"')
    sleep(3)
    cmd = read_yaml()['localfile']['chromedebugwindow']
    win32api.ShellExecute(0, 'open', cmd, '', '', 1)
    sleep(3)
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
        sleep(3)
        top_windows()
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
        browser.switch_to.frame(browser.find_element_by_xpath('//iframe[@class="iframe_panel show"]'))
        el1 = browser.find_element_by_xpath('//li[contains(text(), "税费办理")]')
        ActionChains(browser).move_to_element(el1).perform()
        sleep(0.5)
        taxes_transact2 = browser.find_element_by_xpath('//div[contains(text(), "货物贸易税费支付")]')
        taxes_transact2.click()
        sleep(8)
        browser.switch_to.window(browser.window_handles[1])
        sleep(1)
        tax_index = browser.current_window_handle
        handle_list.append(tax_index)
        relogin_handle_list.append(tax_index)
        tax_should_url = 'https://sz.singlewindow.cn/dyck/swProxy/deskserver/sw/deskIndex?menu_id=spl'
        tax_current_url = browser.current_url
        try:
            if tax_current_url == tax_should_url:
                browser.switch_to.frame(browser.find_element_by_xpath('//iframe[@name="layui-layer-iframe2"]'))
                sleep(1)
                browser.find_element_by_xpath('//img[@alt="关闭"]').click()
                browser.switch_to.default_content()  # 切回默认层
                browser.find_element_by_xpath('//span[text()="支付管理"]').click()
                sleep(1)
                browser.find_element_by_xpath('//a[text()="税费单支付"]').click()
            else:
                #  避免切换海关代码，重新登录后无法进入税单界面
                browser.refresh()
                browser.implicitly_wait(10)
                browser.switch_to.frame(browser.find_element_by_xpath('//iframe[@name="layui-layer-iframe2"]'))
                sleep(1)
                browser.find_element_by_xpath('//img[@alt="关闭"]').click()
                browser.switch_to.default_content()  # 切回默认层
                browser.find_element_by_xpath('//span[text()="支付管理"]').click()
                sleep(1)
                browser.find_element_by_xpath('//a[text()="税费单支付"]').click()
        except:
            browser.refresh()
            browser.implicitly_wait(10)
            browser.switch_to.frame(browser.find_element_by_xpath('//iframe[@name="layui-layer-iframe2"]'))
            sleep(1)
            browser.find_element_by_xpath('//img[@alt="关闭"]').click()
            browser.switch_to.default_content()  # 切回默认层
            browser.find_element_by_xpath('//span[text()="支付管理"]').click()
            sleep(1)
            browser.find_element_by_xpath('//a[text()="税费单支付"]').click()

        #  进入货物申报界面
        browser.switch_to.window(browser.window_handles[0])  # 返回单一首页
        sleep(0.5)
        browser.switch_to.frame(browser.find_element_by_xpath('//iframe[@class="iframe_panel show"]'))
        el2 = browser.find_element_by_xpath('//li[contains(text(), "货物申报")]')
        ActionChains(browser).move_to_element(el2).perform()
        sleep(0.5)
        taxes_transact4 = browser.find_element_by_xpath('//div[contains(text(), "货物申报")]')
        taxes_transact4.click()
        sleep(8)
        browser.switch_to.window(browser.window_handles[-1])
        sleep(1)
        goods_index = browser.current_window_handle
        handle_list.append(goods_index)
        relogin_handle_list.append(goods_index)
        goods_should_url = 'https://sz.singlewindow.cn/dyck/swProxy/deskserver/sw/deskIndex?menu_id=dec001'
        goods_current_url = browser.current_url
        try:
            if goods_current_url == goods_should_url:
                browser.find_element_by_xpath('//span[text()="综合查询"]').click()
                sleep(1)
                browser.find_element_by_xpath('//a[text()= "报关数据查询"]').click()
            else:
                browser.refresh()
                browser.implicitly_wait(10)
                browser.find_element_by_xpath('//span[text()="综合查询"]').click()
                sleep(1)
                browser.find_element_by_xpath('//a[text()= "报关数据查询"]').click()
        except:
            browser.refresh()
            browser.implicitly_wait(10)
            browser.find_element_by_xpath('//span[text()="综合查询"]').click()
            sleep(1)
            browser.find_element_by_xpath('//a[text()= "报关数据查询"]').click()

        browser.service.stop()
        sleep(5)
    # except ElementClickInterceptedException as wl:
    #     tax_ex_content = "检测到网络波动，页面加载失败，请关闭浏览器和爬虫重新操作！"
    #     except_send_email(content=tax_ex_content, ec=wl)
    #     sys.exit(0)
    except BaseException as dl:
        # tax_ex_content = "单一窗口未登录成功，请关闭浏览器和爬虫重新操作！"
        # except_send_email(content=tax_ex_content, ec=dl)
        # sys.exit(0)
        print(dl)


def run_taxe():
    top_windows()
    option1 = ChromeOptions()
    option1.add_experimental_option("debuggerAddress", "127.0.0.1:9000")  # 打开已开启调试窗口
    option1.add_argument('--no-sandbox')
    browser = webdriver.Chrome(options=option1)
    browser.switch_to.window(handle_list[2])  # 切换到货物申报界面
    sleep(2)
    # 切入列表页面iframe
    browser.switch_to.frame(browser.find_element_by_xpath('//iframe[@name="iframe01"]'))
    sleep(1)
    browser.find_element_by_xpath('//input[@id="entryId1"]').clear()
    browser.find_element_by_xpath('//input[@id="entryId1"]').send_keys("524548545812545451")  # 输入待抓取单号
    sleep(1)
    # browser.find_element_by_xpath('//input[@id="operateDate2"]').click()  # 选择本周
    # sleep(1)
    browser.find_element_by_xpath('//button[@id="decQuery"]').click()  # 点击查询

    browser.switch_to.window(handle_list[1])  # 切换到税费window
    browser.switch_to.frame(browser.find_element_by_xpath('//iframe[@name="iframe2"]'))
    sleep(1)
    browser.find_element_by_xpath('//input[@id="entryIdN"]').clear()
    browser.find_element_by_xpath('//input[@id="entryIdN"]').send_keys("524548545812545451")
    browser.find_element_by_xpath('//button[@id="taxationQueryBtnN"]').click()  # 未支付页面查询按钮
    sleep(3)

    browser.switch_to.window(handle_list[2])  # 切换到货物申报界面
    sleep(2)
    # 切入列表页面iframe
    browser.switch_to.frame(browser.find_element_by_xpath('//iframe[@name="iframe01"]'))
    sleep(1)
    browser.find_element_by_xpath('//input[@id="entryId1"]').clear()
    browser.find_element_by_xpath('//input[@id="entryId1"]').send_keys("524548545812545451")  # 输入待抓取单号
    sleep(1)
    # browser.find_element_by_xpath('//input[@id="operateDate2"]').click()  # 选择本周
    # sleep(1)
    browser.find_element_by_xpath('//button[@id="decQuery"]').click()  # 点击查询

    browser.switch_to.window(handle_list[1])  # 切换到税费window
    browser.switch_to.frame(browser.find_element_by_xpath('//iframe[@name="iframe2"]'))
    sleep(1)
    browser.find_element_by_xpath('//input[@id="entryIdN"]').clear()
    browser.find_element_by_xpath('//input[@id="entryIdN"]').send_keys("524548545812545451")
    browser.find_element_by_xpath('//button[@id="taxationQueryBtnN"]').click()  # 未支付页面查询按钮
    sleep(3)


if __name__ == '__main__':
    create_debugwindow()
    sleep(10)
    # run_taxe()



# top_windows()
# option1 = ChromeOptions()
# option1.add_experimental_option("debuggerAddress", "127.0.0.1:9000")  # 打开已开启调试窗口
# option1.add_argument('--no-sandbox')
# browser = webdriver.Chrome(options=option1)
# browser.switch_to.frame(browser.find_element_by_xpath('//iframe[@class="iframe_panel show"]'))
# # taxes_transact2 = browser.find_element_by_xpath('//li[contains(text(), "税费办理")]')
# # print(taxes_transact2)
# # taxes_transact2.click()
# el1 = browser.find_element_by_xpath('//li[contains(text(), "税费办理")]')
# ActionChains(browser).move_to_element(el1).perform()
# sleep(1)
# taxes_transact = browser.find_element_by_xpath('//div[contains(text(), "货物贸易税费支付")]')
# print(taxes_transact)
# taxes_transact.click()
# sleep(1)
# browser.switch_to.window(browser.window_handles[0])
# sleep(2)
# browser.switch_to.frame(browser.find_element_by_xpath('//iframe[@class="iframe_panel show"]'))
# # taxes_transact1 = browser.find_element_by_xpath('//li[contains(text(), "货物申报")]')
# # print(taxes_transact1)
# # taxes_transact1.click()
# el2 = browser.find_element_by_xpath('//li[contains(text(), "货物申报")]')
# ActionChains(browser).move_to_element(el2).perform()
# sleep(1)
# taxes_transact3 = browser.find_element_by_xpath('//div[contains(text(), "货物申报")]')
# print(taxes_transact3)
# taxes_transact3.click()
# sleep(10)


# def publisher(data):
#     """
#     路由模式(direct)
#     """
#     credentials = pika.PlainCredentials(username='admin', password='admin')
#     params = pika.ConnectionParameters(host='192.168.2.80',
#                                        port=5672,
#                                        virtual_host='/',
#                                        credentials=credentials)
#
#     connection = pika.BlockingConnection(params)
#     channel = connection.channel()
#
#     # 声明交换机指定类型 交换机持久化: durable=True, 服务重启后交换机依然存在
#     channel.exchange_declare(exchange='hgBill.autoGetTax.topic.exchange', exchange_type='direct', durable=True)
#
#     properties = pika.BasicProperties(delivery_mode=2)
#
#     # 指定routing_key
#     channel.basic_publish(exchange='hgBill.autoGetTax.topic.exchange',
#                           routing_key='hgBill.autoGetTax',
#                           body=data,
#                           properties=properties)
#
#     connection.close()
#
#
# publisher('BGRQ#123456789123456789')


