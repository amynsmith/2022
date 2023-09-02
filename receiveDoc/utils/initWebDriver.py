import pandas as pd
import datetime as dt
import numpy as np
import os
import glob
import re

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
# from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

# TODO
username = ""
userpass = ""

homehref = "https://..."

# for bulk download file
pagehref_bulkdownload = "https://..."
# for bulk receive file
# templateId changed since 2022.5
pagehref_bulkreceive = "https://..."

# for bulk upload file
basehref_bulkupload = "https://.."

# example revise href
# https://..

# to check
# in chrome devtools, select iframe
# $0.src
sent_pagehref_receive = "https://..."

sent_pagehref_upload = "https://..."

# for processProcedure
# select corresponding tag 'a', then in chrome console, type: $0.baseURI
# receive type 收文待办
todo_pageref_receive = "https://..."
# send type 发文待办
todo_pageref_send = "https://..."
# cooperate type 协同待办
todo_pageref_cooperate = "https://..."

# for bulk download file
basehref_download = "https://..."

basehref_detail = "https://..."
appendhref_detail = "&openFrom=listPending"
basehref_detail_coop = "https://..."

# executable_path = r"C:\Software\ChromeDriver\chromedriver.exe"
executable_path = "chromedriver.exe"  # to freeze project
os.environ["webdriver.chrome.driver"] = executable_path
chrome_options = Options()
# chrome_options.add_extension('path_to_extension')
# profilepath=r"C:\Users\Amy19\AppData\Local\Google\Chrome\User Data\Default"
plugin_loc = r"C:\Program Files (x86)\KGChromePlugin\KGChromePlugin_64.dll;application/kg-plugin"
# chrome_options.add_argument("user-data-dir="+profilepath)
chrome_options.add_argument("register-pepper-plugins=" + plugin_loc)
chrome_options.add_argument("no-sandbox")

# https://stackoverflow.com/questions/16149610/how-to-override-default-set-of-chrome-command-line-switches-in-selenium
# chrome://version/
# https://pywinauto.readthedocs.io/en/latest/getting_started.html#getting-started-guide
# Notes: Chrome requires --force-renderer-accessibility cmd flag before starting.
# Custom properties and controls are not supported because of comtypes Python library restrictions.

chrome_options.add_experimental_option(
    'excludeSwitches',
    ['disable-hang-monitor',
     'disable-prompt-on-repost',
     'disable-background-networking',
     'disable-sync',
     'disable-translate',
     'disable-web-resources',
     'disable-client-side-phishing-detection',
     'disable-component-update',
     'disable-zero-browsers-open-for-tests',
     'force-renderer-accessibility'])

# chrome_options.set_capability('unhandledPromptBehavior', 'accept')
# chrome_options.set_capability("unhandledPromptBehavior", "accept and notify")
chrome_options.set_capability("unhandledPromptBehavior", "ignore")


def init(default_download_path=""):
    if default_download_path != "" and os.path.exists(default_download_path):
        prefs = {
            "download.default_directory": default_download_path,
            "profile.default_content_setting_values.automatic_downloads": 1,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True
        }
        chrome_options.add_experimental_option('prefs', prefs)
    browserservice = Service(executable_path)
    driver = webdriver.Chrome(service=browserservice, options=chrome_options)
    # driver = webdriver.Chrome(executable_path=executable_path, chrome_options=chrome_options)
    driver.maximize_window()  # if to upload file, later set window size
    driver.get(homehref)
    # time wait by seconds
    driver.implicitly_wait(120)  # needs to be called per session
    driver.find_element(By.ID, "login_username").send_keys(username)
    driver.find_element(By.ID, "login_password").send_keys(userpass)
    driver.find_element(By.ID, "login_button").click()
    return driver


# TODO reuse browser session

if __name__ == "__main__":
    default_download_path = r"C:\Users\Amy19\Desktop\收文20220930"
    driver = init(default_download_path)
    size = driver.get_window_size()
    driver.set_window_rect(0, 0, size["width"] * 0.5, size["height"])
    input("...")
