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


def read_yaml():
    try:
        # with open('C:/dist/monitor_downloadtax/configure.yml', encoding="utf-8") as y:
        with open('C:/Users/admin/Desktop/tax_crawling/configure.yml', encoding="utf-8") as y:
            yaml_data = yaml.safe_load(y)
        return yaml_data
    except BaseException as t:
        except_send_email(ec=t)


def except_send_email(ec):
    yag = yagmail.SMTP(user=read_yaml()['email']['user'],
                       password=read_yaml()['email']['password'],
                       host=read_yaml()['email']['host'],
                       port=read_yaml()['email']['port'])
    contents = '监测到grab_tax爬虫程序意外终止，请尝试关闭爬虫程序和浏览器窗口，然后按照使用手册重新运行！\n %s' % ec  # 邮件内容
    subject = 'grab_tax爬虫程序意外终止！'  # 邮件主题
    receiver = read_yaml()['email']['receiver']  # 接收方邮箱账号
    yag.send(receiver, subject, contents)
    yag.close()


def upload_send_email(uc):
    yag = yagmail.SMTP(user=read_yaml()['email']['user'],
                       password=read_yaml()['email']['password'],
                       host=read_yaml()['email']['host'],
                       port=read_yaml()['email']['port'])
    contents = '监测到grab_tax爬虫程序爬取税费文件异常，请尽快核实！\n %s' % uc  # 邮件内容
    subject = 'grab_tax爬虫程序抓取异常！'  # 邮件主题
    receiver = read_yaml()['email']['receiver']  # 接收方邮箱账号
    yag.send(receiver, subject, contents)
    yag.close()


try:
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
except BaseException as p:
    except_send_email(ec=p)

#  创建chrome浏览器调试窗口
cmd = read_yaml()['localfile']['chromedebugwindow']
win32api.ShellExecute(0, 'open', cmd, '', '', 1)


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


def login_taxpage(browser):
    try:
        browser.get('https://sz.singlewindow.cn/dyck/')
        browser.execute_script('() =>{ Object.defineProperties(navigator,{ webdriver:{ get: () => false } }) }')
        close_inform(browser)  # 关闭弹出通知

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
        browser.switch_to.frame(browser.find_element_by_xpath('//div[@id="content"]/iframe'))
        taxes_transact = browser.find_element_by_xpath('//li[contains(text(), "税费办理")]')
        taxes_transact.click()
        sleep(3)
        browser.switch_to.window(browser.window_handles[1])
        should_url = 'https://sz.singlewindow.cn/dyck/swProxy/deskserver/sw/deskIndex?menu_id=spl'
        current_url = browser.current_url
        if current_url == should_url:
            pass
        else:
            ex_content = "监测到程序运行异常，请检查操作员卡是否插好，并按照使用手册重新运行程序！"
            except_send_email(ec=ex_content)
            sys.exit(0)

    except BaseException as o:
        except_send_email(ec=o)


#  启动脚本时登录到税费界面
try:
    option = ChromeOptions()
    option.add_experimental_option("debuggerAddress", "127.0.0.1:9000")  # 打开已开启调试窗口
    option.add_argument('--start-maximized')
    option.add_argument('--no-sandbox')
    browser = webdriver.Chrome(options=option)
    win = browser.window_handles
    win_number = len(win)
    if win_number == 1:
        login_taxpage(browser)
    else:
        pass
except BaseException as s:
    except_send_email(ec=s)

# 连接rabbit
connection = pika.BlockingConnection(pika.ConnectionParameters(host=read_yaml()['rabbitmq']['host'],
                                                               port=read_yaml()['rabbitmq']['port']))
channel = connection.channel()

# 声明direct
channel.exchange_declare(exchange=read_yaml()['rabbitmq']['exchange'], exchange_type='direct', durable=False)

channel.queue_declare(queue=read_yaml()['rabbitmq']['queue'])

queue_num = int(channel.queue_declare(queue=read_yaml()['rabbitmq']['queue']).method.message_count)
# print(queue_num)

# 队列绑定交换机
channel.queue_bind(exchange=read_yaml()['rabbitmq']['exchange'], queue=read_yaml()['rabbitmq']['queue'],
                   routing_key=read_yaml()['rabbitmq']['routing_key'])


def rand_sleep():
    time = random.randint(1, 5)
    # time = random.randint(30, 60)
    return time


def taxdetail_regular_extraction(*args):
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


def analysis_taxfile(file):
    try:
        allgoods_list = []
        taxfilenumber = 0
        with pdfplumber.open(file) as t:
            for page_content in t.pages:
                allfile_content = page_content.extract_text().replace('\n', ' ')

                #  提取税费信息,转为字典
                if '税费单详细信息' in allfile_content:
                    if int(taxfilenumber) == 0:
                        taxbill_detail_dict1 = taxdetail_regular_extraction(allfile_content)
                    elif int(taxfilenumber) == 1:
                        taxbill_detail_dict2 = taxdetail_regular_extraction(allfile_content)
                    taxfilenumber += 1

                #  提取货物信息
                if '税费单货物信息' in allfile_content:
                    detail = re.findall(r"折算率 税率 税率 (.*?)$", allfile_content)
                    strdetail = ''.join(detail)
                    str_detail = strdetail + ' '
                    listdetail = re.findall(r"(.*?) ", str_detail)

                    #  拆分列表，拆分后一个子列表为一行货物明细数据
                    cd = 10
                    if len(listdetail) > cd:
                        for i in range(int(len(listdetail) / cd)):
                            cut_a = listdetail[cd * i:cd * (i + 1)]
                            allgoods_list.append(cut_a)

                        last_data = listdetail[int(len(listdetail) / cd) * cd:]
                        if last_data:
                            allgoods_list.append(last_data)
                    else:
                        allgoods_list.append(listdetail)
        t.close()
        #  解包拆分后的列表，转为字典后再转json
        if int(taxfilenumber) == 1:
            goods_detail_list = []
            for singe_goods in allgoods_list:
                singetax_goodsdict = goodsdetail_regular_extraction(*singe_goods)
                goods_detail_list.append(singetax_goodsdict)
            taxbill_detail_dict1['税费单货物信息'] = goods_detail_list
            taxbill_detail_json = json.dumps(taxbill_detail_dict1, ensure_ascii=False)
            return taxbill_detail_json
        elif int(taxfilenumber) == 2:
            goods_detail_list1 = []
            goods_detail_list2 = []
            singe_taxfile_goodsdetailnumber = int(len(allgoods_list)) / int(taxfilenumber)
            goodsdetaillist1 = allgoods_list[:int(singe_taxfile_goodsdetailnumber)]
            goodsdetaillist2 = allgoods_list[int(singe_taxfile_goodsdetailnumber):]

            for singe_goods in goodsdetaillist1:
                singetax_goodsdict1 = goodsdetail_regular_extraction(*singe_goods)
                goods_detail_list1.append(singetax_goodsdict1)
            taxbill_detail_dict1['税费单货物信息'] = goods_detail_list1
            taxbill_detail_json1 = json.dumps(taxbill_detail_dict1, ensure_ascii=False)

            for singe_goods in goodsdetaillist2:
                singetax_goodsdict2 = goodsdetail_regular_extraction(*singe_goods)
                goods_detail_list2.append(singetax_goodsdict2)
            taxbill_detail_dict2['税费单货物信息'] = goods_detail_list2
            taxbill_detail_json2 = json.dumps(taxbill_detail_dict2, ensure_ascii=False)

            return str([taxbill_detail_json1, taxbill_detail_json2]).replace("'", "")
    except BaseException as v:
        except_send_email(ec=v)


def upload_json():
    try:
        save_path = read_yaml()['localfile']['save_path']
        uploaded_path = read_yaml()['localfile']['uploaded_path']
        uploadfail_path = read_yaml()['localfile']['uploadfail_path']
        filelist = os.listdir(save_path)
        filelist.sort(key=lambda fn: os.path.getmtime(save_path + '\\' + fn))
        name = ''.join(filelist[-1])
        file_name = save_path + name

        url = read_yaml()['upload_api']['api']
        header = {}
        json = analysis_taxfile(file=file_name)
        r = requests.post(url=url, headers=header, json=json)

        code = int(r.json()['code'])
        msg = r.json()['msg']
        if code == 200:
            shutil.move(file_name, uploaded_path)
            print("文件 %s 上传成功！" % name)
        else:
            print("文件 %s 上传失败,原因：%s" % (name, msg))
            uploadfail_content = "税费文件 %s 上传失败，请在 %s 文件中查看!\n失败原因：%s" % (name, uploadfail_path, msg)
            upload_send_email(uc=uploadfail_content)
            shutil.move(file_name, uploadfail_path)
        #  上传接口暂未定义
    except BaseException as u:
        pass

def upload_taxfile():
    try:
        #  文件参数
        save_path = read_yaml()['localfile']['save_path']
        uploaded_path = read_yaml()['localfile']['uploaded_path']
        uploadfail_path = read_yaml()['localfile']['uploadfail_path']
        filelist = os.listdir(save_path)
        filelist.sort(key=lambda fn: os.path.getmtime(save_path + '\\' + fn))
        name = ''.join(filelist[-1])
        file_name = save_path + name
        #  接口参数
        url = read_yaml()['upload_api']['api']
        content = open(file_name, 'rb')
        files = {'file': (name, content, 'pdf')}
        data = {'Content-Disposition': 'form-data', 'Content-Type': 'application/pdf'}
        r = requests.post(url=url, data=data, files=files)
        content.close()
        code = int(r.json()['code'])
        msg = r.json()['msg']
        if code == 200:
            shutil.move(file_name, uploaded_path)
            print("文件 %s 上传成功！" % name)
        else:
            print("文件 %s 上传失败,原因：%s" % (name, msg))
            uploadfail_content = "税费文件 %s 上传失败，请在 %s 文件中查看!\n失败原因：%s" % (name, uploadfail_path, msg)
            upload_send_email(uc=uploadfail_content)
            shutil.move(file_name, uploadfail_path)
    except BaseException as q:
        except_send_email(ec=q)


def go_to_download():
    sleep(2)
    # pyautogui.press('tab', presses=8, interval=0.3)  # 切换到下载
    # sleep(0.5)
    # pyautogui.press('enter')  # 点击下载
    # sleep(1)
    pyautogui.hotkey('ctrl', 's')  # 保存
    sleep(0.5)
    pyautogui.press('left')
    sleep(0.5)
    pyautogui.write(read_yaml()['localfile']['file_path_prefix'])
    sleep(0.8)
    pyautogui.press('shift')  # 防止中文输入法
    sleep(0.8)
    pyautogui.press('enter')  # 点击保存本地
    sleep(0.8)
    pyautogui.hotkey('ctrl', 'w')  # 关闭下载页面
    sleep(1)


# 回调爬取
def callback(ch, method, properties, body):
    try:
        # save_path = read_yaml()['localfile']['save_path']
        # file = os.listdir(save_path)
        # file_num = int(len(file))
        # if file_num > 0:
        #     shutil.rmtree(save_path)
        #     sleep(1)
        #     os.makedirs(save_path)
        taxno = int(body.decode())
        sleep(rand_sleep())
        option = ChromeOptions()
        option.add_experimental_option("debuggerAddress", "127.0.0.1:9000")  # 打开已开启调试窗口
        option.add_argument('--start-maximized')
        option.add_argument('--no-sandbox')
        browser = webdriver.Chrome(options=option)
        if taxno != 123456789:
            browser.switch_to.window(browser.window_handles[0])  # 切换到税费window
            browser.refresh()
            sleep(3)
            browser.switch_to.frame(browser.find_element_by_xpath('//iframe[@name="layui-layer-iframe2"]'))  # 切入iframe
            sleep(1)
            browser.find_element_by_xpath('//img[@alt="关闭"]').click()
            browser.switch_to.default_content()  # 切回默认层
            browser.find_element_by_xpath('//span[text()="支付管理"]').click()
            sleep(0.5)
            browser.find_element_by_xpath('//a[text()="税费单支付"]').click()
            sleep(3)
            browser.switch_to.frame(browser.find_element_by_xpath('//iframe[@name="iframe2"]'))  # 切入列表页面iframe
            sleep(0.5)
            browser.find_element_by_xpath('//input[@id="entryIdN"]').clear()
            browser.find_element_by_xpath('//input[@id="entryIdN"]').send_keys(taxno)
            browser.find_element_by_xpath('//button[@id="taxationQueryBtnN"]').click()  # 未支付页面查询按钮
            sleep(2)
            browser.find_element_by_xpath('(//input[@name="btSelectAll"])[1]').click()  # 勾选数据
            browser.find_element_by_xpath('//button[@id="taxationPrintNButton"]').click()  # 点击预览打印
            sleep(2)
            window = browser.window_handles
            window_number = len(window)
            if window_number == 3:
                go_to_download()
                sleep(10)
                #  可通过修改配置文件切换上传模式：二进制文件流 或 json对象
                upload_mode = read_yaml()['upload_api']['upload_type']
                if int(upload_mode) == 1:
                    upload_taxfile()
                else:
                    upload_json()
                sleep(3)
                ch.basic_ack(delivery_tag=method.delivery_tag)
            else:
                browser.refresh()
                except_content = '报关单 %i 未在单一系统未支付税费模块查询到相关税费文件，请尽快核实' % taxno
                upload_send_email(uc=except_content)
                ch.basic_ack(delivery_tag=method.delivery_tag)
        else:
            browser.switch_to.window(browser.window_handles[0])  # 切换到税费window
            browser.refresh()
            ch.basic_ack(delivery_tag=method.delivery_tag)
    except BaseException as r:
        except_send_email(ec=r)


# 监听队列参数
channel.basic_consume(queue=read_yaml()['rabbitmq']['queue'],
                      auto_ack=False,
                      on_message_callback=callback)
print('正在等待信息，如果想退出，请直接关闭浏览器和程序窗口')

# 启动监听
channel.start_consuming()

if __name__ == "__main__":
    callback()
