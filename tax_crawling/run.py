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
from monitor_downloadtax import *


if __name__ == "__main__":
    # 启动
    create_files()
    create_debugwindow()
    channel.start_consuming()
