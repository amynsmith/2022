import logging.config
import yaml
logger = logging.getLogger('verboseLogger')

import pandas as pd
from datetime import datetime
import numpy as np

import glob
import re
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException,UnexpectedAlertPresentException,ElementClickInterceptedException

import os, sys
import logging
# logging.basicConfig(filename='无oa流程.log', encoding='utf-8', level=logging.DEBUG)
# logging.warning("%s", curfolder.split("\\")[-1])
# logger = logging.getLogger(__name__)

# logger.setLevel(logging.DEBUG)


sys.path.append(os.path.abspath(".."))
import utils.initWebDriver as initdriver
from prepareData.loaddf import loaddf, getinput

# element=driver.find_element(By.CSS_SELECTOR,"table[title='分公司收文文单'] a")
# element.click()

selector_recnum = "#shouxieInput_serial_no"
selector_recdate = "#field0006"

# selector_recsource = "#field0009"
selector_recsource = "#field0009_txt"

selector_doc_mark = "#field0010"
selector_pagenum = "#field0004"
selector_doc_date = "#field0012"
selector_doc_title = "#field0008"
selector_upload_btn = "#uploadGovdocBody_a"
selector_upload_confirm = "#b1"
selector_popup_confirm = "a[title='确定']"
selector_uploadinput = "#fileInputDiv1 input[type='file']"
selector_uploadattach = "#upload input[type='file']"
selector_submit = "#sendId_a"
id_iframe = "govDocZwIframe"
id_iframe_upload = "layui-layer-iframe1"
selector_outerframe = "iframe[id^='layui-layer-iframe']"
selector_popup_close = "a[title='关闭']"
selector_sent_recno = "//*[@id='list']/tr[1]/td[@abbr='serialNo']/div"

dstlist = [selector_recnum, selector_recdate, selector_recsource, \
           selector_doc_mark, selector_pagenum, selector_doc_date, selector_doc_title]
pagehref = initdriver.pagehref_bulkreceive
sentpagehref = initdriver.sent_pagehref_receive


def fillreceiveform(driver: webdriver, inputlist: list):
    logger.info("Filling form...")
    # WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
    driver.switch_to.frame(id_iframe)
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#shouxieIcon_serial_no"))).click()
    for ind, cur_input in enumerate(inputlist):
        driver.find_element(By.CSS_SELECTOR, dstlist[ind]).send_keys(cur_input)
    driver.switch_to.default_content()
    # input("press to confirm...\n")
    return


def uploadreceivefilebody(driver: webdriver, fpath: str):
    mainhandle = driver.current_window_handle
    logger.info("Start uploading body file...")
    driver.find_element(By.CSS_SELECTOR, selector_upload_btn).click()
    # WebDriverWait(driver, 120).until(EC.frame_to_be_available_and_switch_to_it(id_iframe_upload))
    driver.switch_to.frame(id_iframe_upload)
    element = driver.find_element(By.CSS_SELECTOR, selector_uploadinput)
    element.send_keys(fpath)
    # confirm to close upload window
    driver.find_element(By.CSS_SELECTOR, selector_upload_confirm).click()
    # confirm popup warning "重新上传正文"
    driver.switch_to.default_content()
    driver.find_element(By.CSS_SELECTOR, selector_popup_confirm).click()
    # print("may need to manually close file preview window...")
    for cur_handle in driver.window_handles:
        if cur_handle != mainhandle:
            driver.switch_to.window(cur_handle)
            logger.debug(f"closing window {cur_handle}")
            driver.close()
    driver.switch_to.window(mainhandle)
    # input("press to continue...\n")
    return


def uploadreceivefileattach(driver: webdriver, attachpathlist: list):
    if len(attachpathlist) <= 0:
        logger.info("No attach files to upload...")
        return
    logger.info("Start uploading attach files...")
    tmpstr = "\n".join([os.path.basename(tmp) for tmp in attachpathlist])
    logger.info(tmpstr)
    # upload attach files
    element = driver.find_element(By.CSS_SELECTOR, selector_uploadattach)
    attachstr = " \n ".join(attachpathlist)
    element.send_keys(attachstr)
    attachnumstr=str(len(attachpathlist))
    # driver.implicitly_wait(0)  # no mix of two kinds of wait?
    WebDriverWait(driver, 60).until(EC.text_to_be_present_in_element((By.CSS_SELECTOR, "#attachmentNumberDivAtt"),attachnumstr))
    # input("press to confirm...\n")
    return


def receivefile(driver: webdriver, inputlist: list, fpath: str, attachpathlist: list, isthelast: bool, isthefirst: bool):
    # 编号（加括号写年份） 收文日期 来文单位 来文字号 页数 成文日期 标题
    driver.get(pagehref)
    driver.implicitly_wait(120)  # needs to be called per session
    fillreceiveform(driver, inputlist)
    uploadreceivefilebody(driver, fpath)
    uploadreceivefileattach(driver, attachpathlist)
    # submit form
    # driver.find_element(By.CSS_SELECTOR, selector_submit).click()
    # Alert text : 来文文号已被使用，是否继续？
    # TODO adjust for NoAlertPresentException in handling
    # driver.find_element(By.CSS_SELECTOR, selector_submit).click()
    # ElementClickInterceptedException
    # submitclickcount = 0
    # while submitclickcount < 3 and driver.title != "":
    #     try:
    #         driver.find_element(By.CSS_SELECTOR, selector_submit).click()
    #         # WebDriverWait(driver, 30).until(EC.title_is(""))
    #         WebDriverWait(driver, 30).until(EC.url_changes, driver.current_url)
    #     except UnexpectedAlertPresentException:
    #         WebDriverWait(driver, 3).until(EC.alert_is_present())
    #         driver.switch_to.alert.accept()
    #         print("alert accepted")
    #     except TimeoutException:
    #         print("timeout")
    #     submitclickcount = submitclickcount + 1
    #     print(f"submit click count: {submitclickcount}")

    mainhandle = driver.current_window_handle
    try:
        driver.find_element(By.CSS_SELECTOR, selector_submit).click()
        if isthelast or isthefirst:
            logger.info("is the last/first")
            try:
                WebDriverWait(driver, 5).until(EC.title_is(""))
                WebDriverWait(driver, 5).until(EC.url_changes, driver.current_url)
            except TimeoutException:
                logger.info("timeout, click submit twice")
                driver.find_element(By.CSS_SELECTOR, selector_submit).click()
                WebDriverWait(driver, 20).until(EC.title_is(""))
        else:
            WebDriverWait(driver, 30).until(EC.url_changes, driver.current_url)
    except UnexpectedAlertPresentException:
        WebDriverWait(driver, 3).until(EC.alert_is_present())
        driver.switch_to.alert.accept()
        logger.info("alert accepted")

    # TODO ElementClickInterceptedException
    # except ElementClickInterceptedException:
    #     # 选择打开方式（安装或更新PDF插件打开/OFD阅读器兼容打开）
    #     # click "关闭"
    #     id_outerframe = driver.find_element(By.CSS_SELECTOR, selector_outerframe).get_property("id")
    #     driver.switch_to.frame(id_outerframe)
    #     driver.find_element(By.CSS_SELECTOR, selector_popup_close).click()
        ## cannot submit again

    except TimeoutException:
        input("timeout, press to continue...\n")

    driver.switch_to.default_content()
    if isthelast:
        input("press to inspect sent page...\n")
        driver.get(sentpagehref)
        input("...\n")
    return


# window "current"
def receivedone(driver: webdriver, recno: str):
    # $x("//*[@id='list']/tr[1]/td[@abbr='serialNo']/div")[0].innerText
    driver.get(sentpagehref)
    driver.implicitly_wait(120)  # needs to be called per session
    # check recno if is the same
    sent_recno = driver.find_element(By.XPATH, selector_sent_recno).get_property("innerText")
    if recno == sent_recno:
        return True
    else:
        return False


# docmark isattach title recsource fpath pagenum pubdate recno recdate toreceive toupload
def bulkreceivefile(driver: webdriver, fn: str, resfn: str):
    if os.path.exists(resfn):
        logger.info(f"continue where left according to {os.path.basename(resfn)}")
        loadeddf, df = loaddf(resfn, mode="toreceive")
    else:
        logger.info(f"start new bulk receive by {os.path.basename(fn)}")
        loadeddf, df = loaddf(fn, mode="toreceive")
    df["receivedone"] = False
    groupdf = loadeddf.groupby("docmark")
    logger.info(f"total filenum to receive: {len(groupdf)}")
    count=0
    isthelast = False
    isthefirst = False
    for k in groupdf.groups.keys():
        logger.info(f"\ncount {count} ")
        logger.info(f"docmark: {k}")
        curdf = groupdf.get_group(k)
        logger.info(f"title: {curdf[~curdf.isattach].title.values[0]}")
        inputlist, fpath, attachpathlist = getinput(curdf, mode="toreceive")
        logger.info(f"has attach filenum: {len(attachpathlist)}\n")
        if count == len(groupdf)-1:
            isthelast = True
        if count == 0:
            isthefirst = True
        receivefile(driver, inputlist, fpath, attachpathlist, isthelast, isthefirst)
        mainhandle = driver.current_window_handle
        # check if sent success in new tab
        driver.execute_script(f"window.open('about:blank', 'current');")
        driver.switch_to.window("current")
        recno = inputlist[0]
        if receivedone(driver, recno):
            df.loc[curdf.index, "receivedone"] = True
            df.loc[curdf.index, "toreceive"] = False  ############ dynamic modify toreceive
        else:
            logger.info("receive fail")
        driver.switch_to.window(mainhandle)
        df.to_excel(resfn, index=False)
        count=count+1
        isthelast = False
        isthefirst = False
        # proceed = input("continue to process？ type y or n...\n")
        # if proceed == "n":
        #     break
    # df.to_excel(resfn)



if __name__ == "__main__":
    with open('../config.yaml', 'r') as f:
        config = yaml.safe_load(f.read())
        logging.config.dictConfig(config)
    logger = logging.getLogger('verboseLogger')

    # fn = r"C:\Users\Amy19\Desktop\收文20220402\respdflist_modified.xlsx"
    # resfn = r"C:\Users\Amy19\Desktop\收文20220402\res_receivedf.xlsx"
    driver = initdriver.init()

    base_path = r"C:\Users\Amy19\Desktop\收文"
    # basedir = base_path + pd.Timestamp.today().strftime("%Y%m%d")
    basedir = r"C:\Users\Amy19\Desktop\收文20220930"

    fn = basedir + "\\" + "respdflist_modified.xlsx"
    resfn = basedir + "\\" + "res_receivedf.xlsx"
    bulkreceivefile(driver, fn, resfn)

