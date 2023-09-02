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
from selenium.webdriver.common.action_chains import ActionChains
# from selenium.webdriver.support.ui import Select

# TODO
username = ""
userpass = ""
homehref = "http://archives..."


executable_path = r"C:\Software\ChromeDriver\chromedriver.exe"
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
     'disable-zero-browsers-open-for-tests'])

# chrome_options.set_capability('unhandledPromptBehavior', 'accept')
# chrome_options.set_capability("unhandledPromptBehavior", "accept and notify")
chrome_options.set_capability("unhandledPromptBehavior", "ignore")


def init():
    browserservice = Service(executable_path)
    driver = webdriver.Chrome(service=browserservice, options=chrome_options)
    # driver = webdriver.Chrome(executable_path=executable_path, chrome_options=chrome_options)
    driver.maximize_window()  # if to upload file, later set window size
    driver.get(homehref)
    # time wait by seconds
    driver.implicitly_wait(120)  # needs to be called per session
    driver.find_element(By.ID, "j_username").send_keys(username)
    driver.find_element(By.ID, "j_password").send_keys(userpass)
    driver.find_element(By.ID, "submit").click()
    return driver

if __name__ == "__main__":
    driver = init()
    # size = driver.get_window_size()
    # driver.set_window_rect(0, 0, size["width"] * 0.5, size["height"])
    input("...")
