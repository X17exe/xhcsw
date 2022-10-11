from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver
from time import sleep
from selenium.webdriver import ChromeOptions


def brow():
    """
    创建webdriver
    """
    option = ChromeOptions()
    option.add_experimental_option("debuggerAddress", "127.0.0.1:9000")  # 打开已开启调试窗口
    option.add_argument('--start-maximized')
    option.add_argument('--no-sandbox')
    # option.add_argument("--headless")  # 无头模式
    # option.add_argument("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 \
    #     (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36")
    # option.add_experimental_option('excludeSwitches', ['enable-automation', 'load-extension'])  #以调试模式打开窗口
    # 屏蔽保存密码提示
    # prefs = {'profile.default_content_setting_values': {'notifications': 2}}
    # option.add_experimental_option('prefs', prefs)

    return webdriver.Chrome(options=option)


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

        # 判断页面是否加载完成，返回 True or False
        WebDriverWait(browser, 100).until(lambda x: x.find_element_by_xpath('//a[@id="ic_card"]').is_displayed())
        login_button = browser.find_element_by_xpath('//a[@id="ic_card"]')
        login_button.click()
        browser.switch_to.default_content()  # 切回默认层

        browser.switch_to.frame('loginIframe2')
        WebDriverWait(browser, 100).until(lambda x: x.find_element_by_xpath('//input[@id="password"]').is_displayed())
        password_input = browser.find_element_by_xpath('//input[@id="password"]')
        password_input.send_keys('88888888')
        login = browser.find_element_by_id('loginbutton')
        login.click()
        sleep(10)
        browser.switch_to.default_content()
        sleep(1)
        close_inform(browser)  # 关闭弹出通知
        browser.switch_to.frame(browser.find_element_by_xpath('//div[@id="content"]/iframe'))
        taxes_transact = browser.find_element_by_xpath('//li[contains(text(), "税费办理")]')
        taxes_transact.click()
        sleep(1)

    except BaseException as e:
        print(e)


if __name__ == "__main__":
    browser = brow()
    login_taxpage(browser)
