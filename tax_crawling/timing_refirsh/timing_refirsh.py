from selenium import webdriver
from time import sleep
from selenium.webdriver import ChromeOptions
import pika
import random


def timing_refirsh():
    """
    当队列中没有消息时，循环随机刷新税费页面，保持页面cookie活性
    """
    # 连接rabbit
    connection = pika.BlockingConnection(pika.ConnectionParameters(host='192.168.2.80', port=5672))
    channel = connection.channel()

    # 声明direct
    channel.exchange_declare(exchange='hgBill.autoGetTax.topic.exchange', exchange_type='direct', durable=False)

    channel.queue_declare(queue="hgBill.autoGetTax")

    while True:
        sleep(10)
        queue_num = int(channel.queue_declare(queue="hgBill.autoGetTax").method.message_count)
        print(queue_num)
        if queue_num == 0:
            option = ChromeOptions()
            option.add_experimental_option("debuggerAddress", "127.0.0.1:9000")  # 打开已开启调试窗口
            option.add_argument('--start-maximized')
            option.add_argument('--no-sandbox')
            browser = webdriver.Chrome(options=option)
            browser.switch_to.window(browser.window_handles[0])  # 切换到税费window
            refresh_time = random.randint(5, 10)
            sleep(refresh_time)
            browser.refresh()
        else:
            pass


if __name__ == '__main__':
    timing_refirsh()

