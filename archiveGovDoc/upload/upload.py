import pandas as pd
import datetime as dt
import numpy as np
import os
import glob
import re

from selenium import webdriver
from selenium.webdriver.chrome.service import Service

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException

from utils.initWebDriver import init

selector_mainmenu_collect = "//div[starts-with(@id, 'ext-comp')][.//span='接收中心']"

# id_submenu_collect = "Collection_tree"
# selector_submenu_collect = "//*[@id='Collection_tree']//div[contains(@class, 'x-panel-body')]"
# TODO ElementClickInterceptedException <div class="x-panel-body" id="ext-gen93"
selector_submenu_collect = "//*[@id='Collection_tree']//div[contains(@class, 'x-panel-body')]//a[./span]"

# selector_menu1 = "//div[@id='preArchiveEdit_tree']//a[.//span='企业管理类(A)']"
# selector_menu2 = "//div[@id='preArchiveEdit_tree']//a[.//span='企业管理档案']"
# selector_menu3 = "//div[@id='preArchiveEdit_tree']//a[.//span='2021']"

selector_menu1 = "//div[@id='preArchiveEdit_tree']//div[./a/span='企业管理类(A)']"
selector_menu2 = "//div[@id='preArchiveEdit_tree']//div[./a/span='企业管理档案']"
selector_menu3 = "//div[@id='preArchiveEdit_tree']//div[./a/span='2021']"

selector_checkbox = ".//div[@class='x-grid3-row-checker']"
selector_allcheckercontainer = "//*[@id='preArchiveGrid_InnerFilePanelGrid']//td/div[.//div[@class='x-grid3-hd-checker']]"

selector_rows = "//*[@id='preArchiveGrid_InnerFilePanelGrid']//div[contains(@class,'x-grid3-row')][.//table[@class='x-grid3-row-table']]"
selector_cell_archiveno = ".//td//div[contains(text(),'0303.0600-A')]"

selector_uploadbtn = "//button[@type='button'][./text()='附加文件']"
selector_iframe = "//*[@id='preArchiveGrid_SwfUploadPanel_Wind']//iframe"
selector_input = "//nz-upload//input[@type='file']"
# selector_uploadlist = "//nz-upload-list/div[contains(@class, 'ant-upload-list-text-container')]"

tag_uploadlist = "nz-upload-list"
# tag_uploadmsg = "nz-message"
selector_msgcontainer = "//nz-message-container/div[@class='ant-message']"

selector_submit = "//button[.//span[contains(text(),'开始上传')]]"
selector_upload_windowclose = "//div[./span='附加电子文件']/div[contains(@class, 'x-tool-close')]"  # outside iframe


def openmaintab(driver: webdriver):
    element = driver.find_element(By.XPATH, selector_mainmenu_collect)
    if "x-panel-collapsed" in element.get_property("className"):
        element.click()
    # WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, id_submenu_collect))).click()
    # WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, id_submenu_collect))).click()
    # submenu_collect = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, id_submenu_collect)))
    # driver.execute_script("arguments[0].click();", submenu_collect)

    submenu_collect = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, selector_submenu_collect)))
    driver.execute_script("arguments[0].click();", submenu_collect)

    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, selector_menu1))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, selector_menu2))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, selector_menu3))).click()
    return


def clear_allcheck(driver: webdriver):
    allcheckercontainer = driver.find_element(By.XPATH, selector_allcheckercontainer)
    tdcheck = allcheckercontainer.find_element(By.TAG_NAME, "div")

    if "x-grid3-hd-checker-on" in allcheckercontainer.get_property("className"):
        # click once
        tdcheck.click()
    else:
        # dbl click
        action = ActionChains(driver)
        action.double_click(tdcheck)
        action.perform()


# inside iframe
def submitbtn_ready(driver):
    btn_submit = driver.find_element(By.XPATH, selector_submit)
    if btn_submit.get_property("disabled"):
        return False
    else:
        return True

# inside iframe
def upload_complete_perftype(driver):
    uploadlist = driver.find_element(By.TAG_NAME, tag_uploadlist)
    filenum = uploadlist.get_property("childElementCount")
    btn_submit = driver.find_element(By.XPATH, selector_submit)
    # wait for overlaying message to clear
    # msg = driver.find_elements(By.CSS_SELECTOR, tag_uploadmsg)
    msgcontainer = driver.find_element(By.XPATH, selector_msgcontainer)
    msgcount = msgcontainer.get_property("childElementCount")
    if msgcount > 0:
        return False
    elif btn_submit.get_property("disabled") and filenum == 0:
        return True
    else:
        return False


def upload(uploadfn: str):
    uploaddf = pd.read_excel(uploadfn)
    uploaddf = uploaddf.astype({"toupload": bool})
    uploaddf["toupload"] = uploaddf["toupload"].fillna(True)

    driver = init()
    openmaintab(driver)
    # TODO change pagination
    input("Press to start uploading...\n")
    rows = driver.find_elements(By.XPATH, selector_rows)
    rowcount = len(rows)
    print(f"total rows: {rowcount}")

    # loop by row
    for i in range(rowcount):
        selector_currow = selector_rows + "[" + str(i + 1) + "]"
        cur_row = driver.find_element(By.XPATH, selector_currow)
        if "x-grid3-row-selected" not in cur_row.get_property("className"):
            cur_row.find_element(By.XPATH, selector_checkbox).click()
            print(f"\nselected row {i}")
        # cells = cur_row.find_elements(By.TAG_NAME, "td")
        # get archiveno
        # for j in range(len(cells)):
        #     tmp = cells[j].find_element(By.TAG_NAME, "div").get_property("innerText")
        #     if tmp.__contains__("0303.0600-A"):
        #         archiveno = tmp
        #         print(f"archiveno {archiveno}")
        #         break
        cell_archiveno = cur_row.find_element(By.XPATH, selector_cell_archiveno)
        archiveno = cell_archiveno.get_property("innerText")
        print(f"archiveno {archiveno}")
        curdf = uploaddf[uploaddf.archive_no == archiveno][["ftype", "fpath", "toupload"]]
        if len(curdf[curdf.toupload]) == 0:
            print("already uploaded...")
            # unselect current row
            cur_row.find_element(By.XPATH, selector_checkbox).click()
            continue

        # start to upload files
        ftypeset = set(curdf.ftype)
        driver.find_element(By.XPATH, selector_uploadbtn).click()
        WebDriverWait(driver, 60).until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, selector_iframe)))
        # switched to iframe
        dropdown_ftype = driver.find_element(By.TAG_NAME, "nz-select")
        # upload per ftype
        for t in ftypeset:
            print(f"uploading ftype {t}...")
            # select current ftype
            if "ant-select-open" not in dropdown_ftype.get_property("className"):
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable(dropdown_ftype)).click()
            selector_opt = ".//nz-option-item[@title='" + t + "']"
            driver.find_element(By.XPATH, selector_opt).click()
            # ftype selected equals current ftype
            # ftype_selected = dropdown_ftype.find_element(By.TAG_NAME, "nz-select-item").get_property("title")
            fpathstr = "\n".join(curdf[curdf.ftype == t].fpath.values)
            print(fpathstr)
            fileinput = driver.find_element(By.XPATH, selector_input)
            fileinput.send_keys(fpathstr)
            WebDriverWait(driver, 5).until(submitbtn_ready)
            driver.find_element(By.XPATH, selector_submit).click()
            try:
                WebDriverWait(driver, 30, 3).until(upload_complete_perftype)
            except TimeoutException:
                input("submit timeout, press to continue...")

        # close upload window
        driver.switch_to.default_content()
        btn_close = driver.find_element(By.XPATH, selector_upload_windowclose)
        btn_close.click()
        # # uncheck current row
        # cur_row.find_element(By.XPATH, selector_checkbox).click()
        try:
            WebDriverWait(driver, 3).until(EC.staleness_of(cur_row))
        except TimeoutException:
            input("window auto refresh timeout, press to continue...")
        uploaddf.loc[curdf.index, "toupload"] = False
        uploaddf.to_excel(uploadfn, index=False)
    input("All upload finished...\n")


def uploadreview(uploadfn, downloadfn, resfn):
    uploaddf = pd.read_excel(uploadfn)
    uploaddf["filename"] = uploaddf.apply(lambda row: row.archive_no +"_"+ os.path.basename(row.fpath), axis=1)
    uploaddf=uploaddf[["filename","ftype","fpath"]]
    downloaddf=pd.read_excel(downloadfn)
    downloaddf["filename"]=downloaddf["原文路径"].apply(lambda x: os.path.basename(x))
    downloaddf.rename(columns={"文件类别":"ftype"}, inplace=True)
    downloaddf = downloaddf[["filename","ftype"]]
    resdf = uploaddf.merge(downloaddf, on=["filename","ftype"], how="outer", indicator=True)
    resdf.to_excel(resfn)
    return


# 上传完毕后，从档案系统下载卷内电子文件目录进行比对
if __name__ == "__main__":
    uploadfn = r"C:\Users\Amy19\PycharmProjects\archiveGovDoc\upload\toupload.xlsx"
    # upload(uploadfn)
    downloadfn = r"C:\Users\Amy19\Documents\0000-文书归档\2021\2022-06-30 09-18-50\企业管理类电子文件级.xls"
    resfn= r"C:\Users\Amy19\PycharmProjects\archiveGovDoc\upload\uploadreview.xlsx"
    uploadreview(uploadfn, downloadfn, resfn)