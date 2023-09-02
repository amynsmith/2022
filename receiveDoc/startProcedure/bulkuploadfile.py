import logging.config
import yaml
logger = logging.getLogger('verboseLogger')

import pandas as pd
import datetime as dt
import numpy as np
import unicodedata

import glob
import re
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
import subprocess
import pyautogui, pyperclip
import os, sys, time
import logging

sys.path.append(os.path.abspath(".."))
import utils.initWebDriver as initdriver
from prepareData.loaddf import loaddf, getinput
from utils.waitforkeystroke import waitforkeystroke


#把鼠标光标在屏幕左上角，PyAutoGUI函数就会产生pyautogui.FailSafeException异常。
# 如果失控了，需要中断PyAutoGUI函数，就把鼠标光标在屏幕左上角
pyautogui.FAILSAFE = True
#设置成float或int时间（秒），可以为所有的PyAutoGUI函数增加延迟。默认延迟时间是0.1秒
pyautogui.PAUSE = 0.5
recfolder = r"C:\Users\Amy19\PycharmProjects\receiveDoc\resources"
fn_icon = os.path.join(recfolder, "icon_upload.png")
fn_open = os.path.join(recfolder, "icon_open.png")
fn_uploadfinish = os.path.join(recfolder, "icon_pdfuploadfinish.png")
mouseloc_pdfupload=(300,900) #鼠标定位到此处（必须位于主屏幕，且在pdf正文视窗内），然后上传正文pdf
mouseloc_leftcorner = (150,1000) #弹出的打开文件窗口后需要移走鼠标，避免无法定位到弹窗中的打开文件夹图标

basehref = initdriver.basehref_bulkupload
sentpagehref = initdriver.sent_pagehref_upload

id_select = "typeList_id"
id_titleinput = "title"
id_publishbtn = "sendId"
noticetypelist = ["局文件", "其他文件", "X公司文件", "分公司文件"]
selector_checkbox_3 = "//*[@id='openRecordsRead_em']"  # 公开阅读信息
selector_bodytype = "#bodyTypeSelectorspan"
selector_bodytype_pdf = "#menu_bodytype_Pdf"
selector_uploadattach = "*//span[.='添加附件']"
selector_uploadframe = "iframe[id$='-iframe']"
selector_attachinput = "*//input[@type='file']"
selector_upload_confirm = "#b1"
# deprecated, no need to refresh portal
# selector_topnav ="//*[@id='showedNav']//li[@title='XX分公司']"
# selector_allfilebtn = "*//li[@title='所有文件']"
# instead, check uploaded page
selector_uploaded_title = "//*[@id='list']/tr[1]/td[@abbr='o.title']/div"

def filluploadform(driver: webdriver, inputlist: list):
    # inputlist currently includes: title, typeind
    cur_title = inputlist[0]
    cur_noticetype = noticetypelist[inputlist[-1]]
    sel = Select(driver.find_element(By.ID, id_select))
    sel.select_by_visible_text(cur_noticetype)
    # 公开阅读信息
    cur_checkbox = driver.find_element(By.XPATH, selector_checkbox_3)
    if "checked" not in cur_checkbox.get_property("className"):
        cur_checkbox.click()
        logger.info("ensure public reading history")
    else:
        logger.info("public reading history already checked")
    # body format change to pdf
    driver.find_element(By.CSS_SELECTOR, selector_bodytype).click()
    driver.find_element(By.CSS_SELECTOR, selector_bodytype_pdf).click()
    # 若为修改正文，且原公告格式为pdf，则不会出现提示
    try:
        WebDriverWait(driver, 10).until(EC.alert_is_present())
        driver.switch_to.alert.accept()
    except TimeoutException:
        print("No alert of format change")
    # fill title info
    title_elem = driver.find_element(By.ID, id_titleinput)
    title_elem.clear()
    title_elem.send_keys(cur_title)
    ## no need to switch back
    ## driver.switch_to.default_content()
    # input("press to confirm filled upload form...\n")
    return


# def uploadfilebody(driver: webdriver, fpath: str):
#     # mainhandle = driver.current_window_handle
#     logger.info("Start uploading body file, manually drag and drop...")
#     logger.info(fpath)
#     subprocess.run(["explorer.exe", "/select,", os.path.normpath(fpath)])
#     # driver.switch_to.window(mainhandle)
#     input("press to finish drag and drop...\n")
#     return


def uploadfilebody(fpath: str):
    logger.info(fpath)
    # input("Start uploading body file, with UI automation, press to continue...\n")
    print("Start uploading body file, with UI automation")
    # time.sleep(10)
    starttime = time.perf_counter()
    waitforkeystroke()
    x, y = pyautogui.locateCenterOnScreen(fn_icon, grayscale=True)
    pyautogui.click(x, y, button='left')
    print("mouse cursor move to the corner, may speed up searching icon_open")
    if pyautogui.onScreen(mouseloc_leftcorner):
        pyautogui.moveTo(mouseloc_leftcorner)
    x, y = pyautogui.locateCenterOnScreen(fn_open, grayscale=True)
    pyautogui.click(x, y, button='left')
    pyperclip.copy(fpath)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press("enter")
    curtime = time.perf_counter()
    elapsed = curtime - starttime
    print(f"time elapsed: {elapsed}")
    if pyautogui.onScreen(mouseloc_pdfupload):
        pyautogui.moveTo(mouseloc_pdfupload)
    else:
        print("mouse cursor MUST locate on the same screen with browser")
    # input("press to finish uploading body file...\n")
    # input("...\n")  # waitforkeystroke 中启动定时器后，secondcounter 会模拟按下enter键，以取消键盘监听
    # auto decide to continue
    found = pyautogui.locateCenterOnScreen(fn_uploadfinish, minSearchTime=10)
    if found is None:
        input("Failed to upload pdf body file. Manually upload, then press to submit...")
    return



def uploadfileattach(driver: webdriver, attachpathlist: list):
    if len(attachpathlist) <= 0:
        logger.info("No attach files to upload...")
        return
    logger.info("Start uploading attach files listed below...")
    tmpstr = "\n".join([os.path.basename(tmp) for tmp in attachpathlist])
    logger.info(tmpstr)
    # upload attach files
    driver.find_element(By.XPATH, selector_uploadattach).click()
    id_uploadframe = driver.find_element(By.CSS_SELECTOR, selector_uploadframe).get_property("id")
    driver.switch_to.frame(id_uploadframe)
    attachstr = " \n ".join(attachpathlist)
    driver.find_element(By.XPATH, selector_attachinput).send_keys(attachstr)
    # input("press to submit attach file...\n")
    driver.find_element(By.CSS_SELECTOR, selector_upload_confirm).click()
    attachnumstr = str(len(attachpathlist))
    driver.switch_to.default_content()
    WebDriverWait(driver, 30).until(EC.text_to_be_present_in_element((By.CSS_SELECTOR, "#attachmentNumberDiv"), attachnumstr))
    return


def uploadfile(driver: webdriver, inputlist: list, fpath: str, attachpathlist: list):
    # 编号（加括号写年份） 收文日期 来文单位 来文字号 页数 成文日期 标题
    driver.get(basehref)
    driver.implicitly_wait(120)  # needs to be called per session
    filluploadform(driver, inputlist)
    # upload attach files automatically
    # then manually drag and drop body file
    # TODO dispatch customized event for external pdf plugin, avoid manual upload of body file
    uploadfileattach(driver, attachpathlist)
    # uploadfilebody(driver, fpath)
    uploadfilebody(fpath)
    # submit form
    # jump = input("LAST, press j\t to jump, else defaults to publish...\n")
    # if jump == "j":
    #     print("quit publishing current file")
    #     driver.close()
    #     return
    # else:
    #     driver.find_element(By.ID, id_publishbtn).click()
    driver.find_element(By.ID, id_publishbtn).click()
    # driver.switch_to.default_content() # window "current" automatically closed



# def bulkuploadfile(driver: webdriver, fn: str, resfn: str):
#     mainhandle = driver.current_window_handle
#
#     if os.path.exists(resfn):
#         logger.info(f"continue where left according to {os.path.basename(resfn)}")
#         loadeddf, df = loaddf(resfn, mode="toupload")
#     else:
#         logger.info(f"start new bulk upload by {os.path.basename(fn)}")
#         loadeddf, df = loaddf(fn, mode="toupload")
#
#     df["uploaddone"] = False
#     groupdf = loadeddf.groupby("docmark")
#     logger.info(f"total filenum to upload: {len(groupdf)}")
#     count = 0
#     for k in groupdf.groups.keys():
#         logger.info(f"\ncount {count}")
#         logger.info(f"docmark: {k}")
#         curdf = groupdf.get_group(k)
#         inputlist, fpath, attachpathlist = getinput(curdf, mode="toupload")
#         logger.info(f"title: {curdf[~curdf.isattach].title.values[0]}")
#         logger.info(f"has attach filenum: {len(attachpathlist)}")
#         # start uploading in new tab
#         driver.execute_script(f"window.open('about:blank', 'current');")
#         driver.switch_to.window("current")
#         uploadfile(driver, inputlist, fpath, attachpathlist)
#         driver.switch_to.window(mainhandle)
#         # finished uploading, switched back to update all files list in portal
#         driver.find_element(By.XPATH, selector_topnav).click()
#         # StaleElementReferenceException
#         # 10 -> 30s
#         WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, selector_allfilebtn))).click()
#         # driver.find_element(By.XPATH, selector_allfilebtn).click()  # deprecated!!! 分公司门户被改动
#         # check if upload success
#         # ################# incomplete
#         df.loc[curdf.index, "uploaddone"] = True
#         df.loc[curdf.index, "toupload"] = False  ############ dynamic modify toreceive
#         df.to_excel(resfn, index=False)
#         count = count + 1
#         # proceed = input("continue to process？ type y or n...\n")
#         # if proceed == "n":
#         #     break
#         WebDriverWait(driver, 30).until(EC.number_of_windows_to_be(1))
#     # df.to_excel(resfn)


# window main, instead of window "current"
def uploaddone(driver: webdriver, title: str):
    driver.get(sentpagehref)
    driver.implicitly_wait(120)  # needs to be called per session
    driver.refresh()
    # check title if is the same
    sent_title = driver.find_element(By.XPATH, selector_uploaded_title).get_property("title")
    sent_title = unicodedata.normalize("NFKD", sent_title)   # eg: 'XX工函〔2022〕29号\xa0关于进一步做好保障农民工工资支付工作的通知'
    title = unicodedata.normalize("NFKD", title)
    if title == sent_title:
        return True
    else:
        return False



def bulkuploadfile(driver: webdriver, fn: str, resfn: str):
    mainhandle = driver.current_window_handle
    if os.path.exists(resfn):
        logger.info(f"continue where left according to {os.path.basename(resfn)}")
        loadeddf, df = loaddf(resfn, mode="toupload")
    else:
        logger.info(f"start new bulk upload by {os.path.basename(fn)}")
        loadeddf, df = loaddf(fn, mode="toupload")

    df["uploaddone"] = False
    groupdf = loadeddf.groupby("docmark")
    logger.info(f"total filenum to upload: {len(groupdf)}")
    count = 0
    for k in groupdf.groups.keys():
        logger.info(f"\ncount {count}")
        logger.info(f"docmark: {k}")
        curdf = groupdf.get_group(k)
        inputlist, fpath, attachpathlist = getinput(curdf, mode="toupload")
        logger.info(f"title: {curdf[~curdf.isattach].title.values[0]}")
        logger.info(f"has attach filenum: {len(attachpathlist)}")
        # upload in window "current"
        driver.execute_script(f"window.open('about:blank', 'current');")
        driver.switch_to.window("current")
        uploadfile(driver, inputlist, fpath, attachpathlist)
        # check if upload success in main window
        # window "current" will auto close, thus only main window remains
        WebDriverWait(driver, 30).until(EC.number_of_windows_to_be(1))
        title = inputlist[0]
        driver.switch_to.window(mainhandle)
        if uploaddone(driver, title):
            df.loc[curdf.index, "uploaddone"] = True
            df.loc[curdf.index, "toupload"] = False  ############ dynamic modify toupload
        else:
            logger.info("upload fail")
        df.to_excel(resfn, index=False)
        count = count + 1




if __name__ == "__main__":
    with open('../config.yaml', 'r') as f:
        config = yaml.safe_load(f.read())
        logging.config.dictConfig(config)
    logger = logging.getLogger('verboseLogger')
    # fn = r"C:\Users\Amy19\Desktop\收文20220402\respdflist_modified.xlsx"
    # resfn = r"C:\Users\Amy19\Desktop\收文20220402\res_uploaddf.xlsx"
    driver = initdriver.init()

    size = driver.get_window_size()
    driver.set_window_rect(0, 0, size["width"] * 0.5, size["height"])
    base_path = r"C:\Users\Amy19\Desktop\收文"
    basedir = base_path + pd.Timestamp.today().strftime("%Y%m%d")
    # basedir = r"C:\Users\Amy19\Desktop\收文20220509"
    fn = basedir + "\\" + "respdflist_modified.xlsx"
    resfn = basedir + "\\" + "res_uploaddf.xlsx"
    bulkuploadfile(driver, fn, resfn)


