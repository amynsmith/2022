import pandas as pd
import datetime as dt
import numpy as np

import glob
import re
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import shutil
import os, sys
import time

sys.path.append(os.path.abspath(".."))
import utils.initWebDriver as initdriver
from prepareData.loaddf import loaddf, getinput

exclude_ptn = re.compile(r"(^res)|(_resave.pdf$)|(.crdownload$)|(.tmp$)")

pagehref = initdriver.pagehref_bulkdownload
basehref_download = initdriver.basehref_download
selector_info = "#listData div.list_infos input"
selector_last = "div#noprint span.left > span:last-of-type"
selector_title = "#head_height div.mainText_head_title div"
id_bodydownlaod = "downloadOriginalContent"
id_attachdownload = "downloadOriginalAttr"


# switch from current window to download page
def chrome_download_complete(driver: webdriver):
    if not driver.current_url.startswith("chrome://downloads"):
        driver.execute_script(f"window.open('about:blank', 'download');")
        driver.switch_to.window("download")
        driver.get("chrome://downloads/")
        # driver.implicitly_wait(10)  # needs to be called per session
    # return driver.execute_script("""
    #     var items = document.querySelector('downloads-manager')
    #         .shadowRoot.getElementById('downloadsList').items;
    #     if (items.every(e => e.state === "COMPLETE"))
    #         return items.map(e => e.fileUrl || e.file_url);
    #     """)
    return driver.execute_script("""
        var items = document.querySelector('downloads-manager')
            .shadowRoot.getElementById('downloadsList').items;
        if (items.every(e => e.state === "COMPLETE"))
            items = items.filter(e => e.state === "COMPLETE" && e.fileExternallyRemoved === false)
            return items.map(e => e.fileUrl || e.file_url);
        """)


def waitfor_download(driver: webdriver, totalnum: int):
    if totalnum == 1:
        WebDriverWait(driver, 120, 1).until(chrome_download_complete)
    else:
        # time.process_time()
        starttime = time.perf_counter()
        max_waittime = 60
        while True:
            paths = WebDriverWait(driver, 120, 2).until(chrome_download_complete)
            if len(paths) >= totalnum:
                print("browser download completed")
                # file name all encoded instead of chinese
                # print("downloaded files:")
                # for i in paths:
                #     print(os.path.basename(i))
                return True
            else:
                curtime = time.perf_counter()
                elapsed = curtime - starttime
                print(f"time elapsed: {elapsed}")
                if elapsed >= max_waittime:
                    print("Warning: browser download incomplete")
                    return False


def moveto_subfolder(download_path, subfoldername, totalnum):
    # check if is file type, for exception:
    rootpath = download_path + r"\*.*"
    subfolder = os.path.join(download_path, subfoldername)
    if not os.path.exists(subfolder):
        os.mkdir(subfolder)
    # for i in glob.iglob(rootpath):
    #     if exclude_ptn.search(os.path.basename(i)) is not None:
    #         print(f"exclude {os.path.basename(i)}")
    #         continue
    #     target = os.path.join(subfolder, os.path.basename(i))
    #     shutil.move(i, target)
    #     print(f"moving file {os.path.basename(i)}")
    tmpl = list(glob.iglob(rootpath))
    srcl = [i for i in tmpl if (os.path.isfile(i) and exclude_ptn.search(os.path.basename(i)) is None)]
    loop_wait = 0
    while len(srcl) != totalnum:
        print(f"current srcl:")
        [print(os.path.basename(i)) for i in srcl]
        print("wait for explorer files to refresh")
        time.sleep(2)
        tmpl = list(glob.iglob(rootpath))
        srcl = [i for i in tmpl if exclude_ptn.search(os.path.basename(i)) is None]
        loop_wait = loop_wait+1
        print(f"current loop_wait: {loop_wait}")
        if loop_wait > 5:
            input("Warn: possible failure of file integrity, manually check and complete download, then press to continue...\n")
            tmpl = list(glob.iglob(rootpath))
            srcl = [i for i in tmpl if exclude_ptn.search(os.path.basename(i)) is None]
            print(f"updated srcl:")
            [print(os.path.basename(i)) for i in srcl]
    for i in srcl:
        target = os.path.join(subfolder, os.path.basename(i))
        shutil.move(i, target)
        print(f"moving file {os.path.basename(i)}")



def get_hreflist(driver, jumpcount):
    elements = driver.find_elements(By.CSS_SELECTOR, selector_info)
    ids = [e.get_property("value") for e in elements]
    if jumpcount > 0:
        ids = ids[:-int(jumpcount)]
    hreflist = [basehref_download + i for i in ids]
    print(f"已跳过 {jumpcount} 行")
    print(f"当前页面需查看文件数量： {len(ids)}")
    return hreflist


# roughly classify into folders created and named with title
def downloadfile(driver: webdriver, download_path: str):
    title = driver.find_element(By.CSS_SELECTOR, selector_title).get_property("innerText")
    element = driver.find_element(By.ID, id_bodydownlaod)
    element.click()
    # not last-child
    last = driver.find_element(By.CSS_SELECTOR, selector_last)
    if last.get_property("id") == id_bodydownlaod:
        totalnum = 1
        # print("无附件")
    else:
        attach_tables = driver.find_elements(By.CSS_SELECTOR, "table.attachment_operate_btn")
        attach_titlelist = [t.get_property("title") for t in attach_tables]
        totalnum = 1 + len(attach_titlelist)
        # for t in attach_titlelist:
        #     print(t)
        element = driver.find_element(By.ID, id_attachdownload)
        element.click()
    print(f"title: {title}")
    print(f"total file num: {totalnum}")
    curhandle = driver.current_window_handle
    # input("press to finish downloading\n")
    # waits for all the files to be completed and returns the paths
    # class selenium.webdriver.support.wait.WebDriverWait(driver, timeout, poll_frequency=0.5, ignored_exceptions=None)
    waitfor_download(driver, totalnum)
    driver.close()
    driver.switch_to.window(curhandle)
    subfoldername = re.sub("\W", "", title)
    moveto_subfolder(download_path, subfoldername, totalnum)
    return subfoldername


def bulkdownloadfile(driver: webdriver, download_path: str, resfn: str):
    driver.get(pagehref)
    driver.implicitly_wait(120)  # needs to be called per session
    mainhandle = driver.current_window_handle
    dfdata = []
    print("First, manually scroll down pages")
    jumpstr = input("type num to jump lines, defaults to 0\n")
    jumpstr = re.sub("\D", "", jumpstr)
    if jumpstr == "":
        jumpcount = 0
    else:
        jumpcount = int(jumpstr)
    hreflist = get_hreflist(driver, jumpcount)
    start_ind = 0
    while True:
        print("\nind number [out of range] leads to terminating bulk download")
        print(f"START_ind/LEN: \t {start_ind}/{len(hreflist)}")
        tmp = input(f"type ind number\t to proceed, defaults to {start_ind - 1}...\n")
        tmp = re.sub("[^-\d]", "", tmp)
        if tmp == "":
            ind = start_ind - 1  # defaults to process in time order
            if ind not in list(range(-len(hreflist), len(hreflist))):
                print("by default terminated bulk download")
                break
        elif int(tmp) not in list(range(-len(hreflist), len(hreflist))):
            print("user terminated bulk download")
            break
        else:
            ind = int(tmp)  # process the specified one
        start_ind = ind
        driver.execute_script(f"window.open('about:blank', 'current');")
        driver.switch_to.window("current")
        driver.get(hreflist[ind])
        driver.implicitly_wait(120)  # needs to be called per session
        skip = input("press s\t to skip current file, else start download\n")
        if skip == "s":
            driver.close()
            driver.switch_to.window(mainhandle)
            continue
        else:
            subfoldername = downloadfile(driver, download_path)
            driver.close()
            driver.switch_to.window(mainhandle)
            dfdata = dfdata + [subfoldername]
    df = pd.DataFrame(data=dfdata, columns=["subfoldername"])
    df.to_excel(resfn)
    return df


def main(download_path, resfn):
    if not os.path.exists(download_path):
        os.mkdir(download_path)
        print("default_download_path now created")
    else:
        print("default_download_path already exists")
    driver = initdriver.init(download_path)
    bulkdownloadfile(driver, download_path, resfn)


if __name__ == "__main__":
    base_path = r"C:\Users\Amy19\Desktop\收文"
    download_path = base_path + pd.Timestamp.today().strftime("%Y%m%d")
    resfn = download_path + "\\" + "res_subfoldername.xlsx"
    main(download_path, resfn)
