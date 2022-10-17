from monitor_downloadtax import *
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

if __name__ == "__main__":
    # 启动
    create_files()
    create_debugwindow()
    channel.start_consuming()
