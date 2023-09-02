import logging.config
import yaml
logger = logging.getLogger('verboseLogger')

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
from selenium.webdriver.support.ui import Select
import subprocess
import os, sys
import logging

sys.path.append(os.path.abspath(".."))
import utils.initWebDriver as initdriver
from utils.constants import people_depart
from startProcedure.bulkuploadfile import filluploadform,uploadfileattach,uploadfilebody
from processProcedure.advancedoperation import selectpeople_by_search,selectpeople_by_branch

basehref = initdriver.basehref_bulkupload
id_select = "typeList_id"
id_titleinput = "title"
id_publishbtn = "sendId"
# id_viewrange = "issueAreaName"
noticetypelist = ["局文件", "其他文件", "X公司文件", "分公司文件"]
selector_checkbox_3 = "//*[@id='openRecordsRead_em']"  # 公开阅读信息
selector_bodytype = "#bodyTypeSelectorspan"
selector_bodytype_pdf = "#menu_bodytype_Pdf"
selector_uploadattach = "*//span[.='添加附件']"
selector_outerframe = "iframe[id$='-iframe']"  # also for upload attach
selector_innerframe = "iframe[id^='layui-layer-iframe']"  # also useful for selecting people
selector_attachinput = "*//input[@type='file']"
selector_upload_confirm = "#b1"
selector_topnav ="//*[@id='showedNav']//li[@title='X分公司']"
selector_allfilebtn = "*//li[@title='所有文件']"

selector_deleteitem = "//*[@id='selectPeopleTable']//span[@title='删除']"

orgtreelist = ["X分公司",
               "X分公司",
               "Y经理部",
               "Z经理部"]

displaytablist = ["部门", "组"]
tabname = displaytablist[0]


def setviewrange(driver: webdriver, inputlist:list):
    # inputlist = [cur_title, namestr, branchlist, cur_noticetype]
    nameslist = people_depart.getnames(inputlist[1])
    print(nameslist)
    branchdict = dict(zip(orgtreelist, inputlist[2]))
    [print(k + "-" + str(branchdict[k])) for k in branchdict]

    # 选择发布范围
    driver.find_element(By.ID, "issueAreaName").click()
    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "spAccount")))

    id_outerframe = driver.find_element(By.CSS_SELECTOR, selector_outerframe).get_property("id")
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, id_outerframe)))
    id_innerframe = driver.find_element(By.CSS_SELECTOR, selector_innerframe).get_property("id")
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, id_innerframe)))

    sel_element = driver.find_element(By.ID, "List3")
    sel = Select(sel_element)
    optnum = sel_element.get_property("childElementCount")  # or length?
    if optnum>0:
        for i in range(optnum):
            sel.select_by_index(i)
        driver.find_element(By.XPATH, selector_deleteitem).click()
        print("default selection cleared")

    print("stage1: selecting by individual name...")
    if len(nameslist)<=0:
        print("skip")
    else:
        # tabname = displaytablist[0]
        selectpeople_by_search(driver, nameslist)

    print("stage2: selecting by branch...")
    if len(branchlist)<=0:
        print("skip")
    else:
        selectpeople_by_branch(driver, branchlist)

    input("LAST, press to finish...\n")
    driver.switch_to.default_content()
    driver.switch_to.frame(id_outerframe)
    driver.find_element(By.CSS_SELECTOR, "a[title='确定']").click()  # beneath outerframe
    driver.switch_to.default_content()
    return





def restrictedupload(driver: webdriver, inputlist: list, fpath: str, attachpathlist: list, revisehref=""):
    mainhandle = driver.current_window_handle
    # start uploading in new tab
    driver.execute_script(f"window.open('about:blank', 'current');")
    driver.switch_to.window("current")

    targethref=[basehref if revisehref == "" else revisehref][0]
    driver.get(targethref)

    driver.implicitly_wait(120)  # needs to be called per session
    filluploadform(driver, inputlist)

    input("press to set view range...\n")
    setviewrange(driver, inputlist)

    uploadfileattach(driver, attachpathlist)
    # uploadfilebody(driver, fpath)
    uploadfilebody(fpath)
    driver.find_element(By.ID, id_publishbtn).click()

    driver.switch_to.window(mainhandle)
    # finished uploading, switched back to update all files list in portal
    driver.find_element(By.XPATH, selector_topnav).click()
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, selector_allfilebtn))).click()
    WebDriverWait(driver, 30).until(EC.number_of_windows_to_be(1))
    return


if __name__ == "__main__":
    # revise
    # 党委会会议纪要
    name1 = """XX XX XX
    YY YY YY
    """  #项目经理

    name2 = """XXX
    YYY"""  #经理部（XX分公司）部门负责人

    tmpl=[people_depart[k].value for k in people_depart.__members__.keys()]
    name3= "XX "+" ".join(tmpl)  # 总部部门负责人
    namestr = name1 + " " + name2+ " " + name3

    branchlist = ["领导班子 总经理助理", "领导班子", "XX经理部领导", "领导班子 总经理助理"]  #领导班子及总助
    fpath=r"C:\Users\Amy19\Documents\00000-收发文2022\发文\10 党委纪要\第10号\第10号 XX党委会会议纪要.pdf"

    cur_noticetype = 3  # 分公司文件
    attachpathlist = []
    # # revisehref = "https://..."
    revisehref=""

    cur_title = os.path.splitext(os.path.basename(fpath))[0]
    inputlist = [cur_title, namestr, branchlist, cur_noticetype]
    driver = initdriver.init()
    restrictedupload(driver, inputlist, fpath, attachpathlist, revisehref)