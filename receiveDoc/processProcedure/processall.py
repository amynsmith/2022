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

import os, sys

sys.path.append(os.path.abspath(".."))
import utils.initWebDriver as initdriver
from utils.constants import getallnames
from processProcedure.advancedoperation import advanceforward

id_iframe = "main"  # todolist page frame
selector_totalnum = "span[id$='_grid_total_number']"  # 列表界面下方显示记录总数
# id_totalnum_coop ="#allPendingNum" #协同类待办列表界面，上方也显示数字
selector_totalnum_coop = "span[id$='_grid_total_number']"  # 列表界面下方显示记录总数

selector_todolist_checkbox = "#list span.ctpUiCheckbox input[type='checkbox']"
# for archive
id_archivecheck = "pigeonhole"
selector_outerframe = "iframe[id^='layui-layer-iframe']"  # also useful for selecting people
id_innerframe = "treeMoveFrame"
selector_docarchive = "#webfx-tree-object-15-anchor"
selector_receivearchive = "#webfx-tree-object-21-anchor"
# #webfx-tree-object-27-anchor #党委
# #webfx-tree-object-29-anchor #团委
# #webfx-tree-object-31-anchor #纪委
# #webfx-tree-object-33-anchor #行政
# #webfx-tree-object-35-anchor #工会
selector_category_dict = {"a": "#webfx-tree-object-27-anchor",
                          "b": "#webfx-tree-object-29-anchor",
                          "c": "#webfx-tree-object-31-anchor",
                          "d": "#webfx-tree-object-33-anchor",
                          "e": "#webfx-tree-object-35-anchor"}
selector_sendarchive = "#webfx-tree-object-17-anchor"  # XX分公司发文2022

# 协同、发文、收文
todo_pageref_dict = {"C": initdriver.todo_pageref_cooperate,
                     "S": initdriver.todo_pageref_send,
                     "R": initdriver.todo_pageref_receive}

# for type Receive
selector_btn_leaders = "#_dealDiv input[value^='领导批示（并）']"
selector_btn_departs = "#_dealDiv input[value^='部门承办（并）']"
# for type Send
# selector_btn_department = "#_dealDiv input[value='分公司部门会签']"
# selector_btn_department_alt = "#_dealDiv input[value='分公司承办']"
selector_btn_department = "input:not([value*='领导'])[value^='分公司']"

selector_btn_midleader = "#_dealDiv input[value='分公司分管领导']"
selector_btn_topleader = "#_dealDiv input[value='分公司主职领导']"
######## 套红头发文
selector_btn_departs_govdoc = "#_dealDiv input[value='分公司承办']"
selector_btn_leaders_govdoc = "#_dealDiv input[value='主管领导审批（并）'"



selector_btn_dict = {"C": [""],
                     "S": [selector_btn_department, selector_btn_midleader, selector_btn_topleader,
                           selector_btn_departs_govdoc, selector_btn_leaders_govdoc],
                     "R": [selector_btn_leaders, selector_btn_departs]}


# common use
def get_hreflist(driver, processtype):
    if processtype in ["S", "R"]:
        driver.switch_to.frame(id_iframe)
        totalnumstr = driver.find_element(By.CSS_SELECTOR, selector_totalnum).get_property("innerText")
    else:
        #processtype == "C"
        totalnumstr = driver.find_element(By.CSS_SELECTOR, selector_totalnum_coop).get_property("innerText")
    totalnum = int(re.sub(r"\D", "", totalnumstr))
    if totalnum == 0:
        print("当前页面待办个数： 0")
        return []
    checkboxes = driver.find_elements(By.CSS_SELECTOR, selector_todolist_checkbox)
    ids = [i.get_property("value") for i in checkboxes]
    print(f"当前页面待办个数： {len(ids)}")
    if processtype in ["S", "R"]:
        hreflist = [initdriver.basehref_detail + tmpid + initdriver.appendhref_detail for tmpid in ids]
    else:
        # processtype == "C":
        hreflist = [initdriver.basehref_detail_coop + tmpid for tmpid in ids]
    return hreflist


# already switched to window "current"
# click and submit
def click_peoplebtn(driver, choice, names=None, processtype="S"):
    cur_selector_btn = selector_btn_dict[processtype][int(choice)]
    driver.find_element(By.CSS_SELECTOR, cur_selector_btn).click()
    id_layoutframe = driver.find_element(By.CSS_SELECTOR, selector_outerframe).get_property("id")
    driver.switch_to.frame(id_layoutframe)

    if processtype == "S":
        if choice == "0" or choice == "3":
            selrange = "depart"
        else:
            selrange = "leader"
    else:
        # processtype == "R"
        if choice == "0":
            selrange = "leader"
        else:
            selrange = "depart"
    if names is None:
        names = getallnames(selrange)

    for n in names:
        element = driver.find_element(By.XPATH,
                                      "//*[@id='selectPeopleBody']//tr[.//div='" + n + "']//td[@class='checkBoxTh']//input[@type='checkbox']")
        print(n)
        element.click()
    if names is None:
        input("press to submit...\n")
    # else auto confirm
    driver.switch_to.default_content()
    driver.find_element(By.CSS_SELECTOR, "a[title='确定']").click()


# already switched to window "current"
# click and submit
def click_archivebtn(driver, category="", processtype="R"):
    driver.find_element(By.ID, id_archivecheck).click()
    driver.find_element(By.CSS_SELECTOR, "input[value='提交']").click()
    id_outerframe = driver.find_element(By.CSS_SELECTOR, selector_outerframe).get_property("id")
    driver.switch_to.frame(id_outerframe)
    driver.switch_to.frame(id_innerframe)
    driver.find_element(By.CSS_SELECTOR, selector_docarchive).click()
    # select corresponding folder
    if processtype == "S":
        driver.find_element(By.CSS_SELECTOR, selector_sendarchive).click()
    elif processtype == "R":
        driver.find_element(By.CSS_SELECTOR, selector_receivearchive).click()
        if category == "":
            print("press A or a\t to choose folder 党委")
            print("press B or b\t to choose folder 团委")
            print("press C or c\t to choose folder 纪委")
            print("press D or d\t to choose folder 行政")
            print("press E or e\t to choose folder 工会")
            category = input("type category number...\n")
        category = category.lower()
        if category in selector_category_dict.keys():
            print(f"flag selected for archive category: {category}")
            selector_category = selector_category_dict[category]
        else:
            print("category defaults to 行政...\n")
            selector_category = selector_category_dict["d"]
        driver.find_element(By.CSS_SELECTOR, selector_category).click()
    # submit
    driver.switch_to.default_content()
    driver.switch_to.frame(id_outerframe)
    driver.find_element(By.CSS_SELECTOR, "#btnRow input[value='确定']").click()


def click_submitbtn(driver):
    driver.find_element(By.CSS_SELECTOR, "input[value='提交']").click()
    return



def processtypeC(driver: webdriver):
    input("manually process, then press to continue...\n")
    return


def processtypeS(driver: webdriver):
    # TODO directly edit processmap to bulk add people and modify nodes
    print("========== for regular requests ==========")
    print("press 0\t to choose department")
    print("press 1\t to choose midleader")
    print("press 2\t to choose topleader")
    print("========== for official document ==========")
    print("press 3\t to choose depart")
    print("press 4\t to choose leader")
    print("press 5\t to choose archive folder")
    print("========== universal ==========")
    print("press 90\t to submit forward")
    print("press 99\t to advance forward")
    choice = input("type choice number...\n")
    if choice in ["0", "1", "2", "3", "4"]:
        # choose people
        click_peoplebtn(driver, choice, processtype="S")
    elif choice == "5":
        click_archivebtn(driver, processtype="S")
    elif choice == "90":
        click_submitbtn(driver)
    elif choice == "99":
        advanceforward(driver, processtype="S")
    else:
        input("manually process, then press to continue...\n")
        return


# after switching to window "current"
def processtypeR(driver: webdriver):
    print("press 0\t to choose leaders")
    print("press 1\t to choose departments")
    print("press 2\t to choose archive folder")
    print("press 90\t to submit forward")
    print("press 99\t to advance forward")
    choice = input("type choice number...\n")
    if choice in ["0", "1"]:
        click_peoplebtn(driver, choice, processtype="R")
    elif choice == "2":
        click_archivebtn(driver, processtype="R")
        # driver.switch_to.window(mainhandle)
    elif choice == "90":
        click_submitbtn(driver)
    elif choice == "99":
        advanceforward(driver, processtype="R")
    else:
        # driver.close()
        # manually process, then close
        input("manually process, then press to continue...\n")
        # driver.switch_to.window(mainhandle)
        return


def processallbytype(driver: webdriver, processtype: str):
    pagehref = todo_pageref_dict[processtype]
    driver.get(pagehref)
    driver.implicitly_wait(120)  # needs to be called per session
    mainhandle = driver.current_window_handle
    hreflist = get_hreflist(driver, processtype)
    start_ind = -1  # defaults to process in time order
    while len(hreflist) > 0:
        print(f"ind number [out of range] leads to terminating this processtype {processtype}")
        tmp = input(f"type ind number\t to proceed, defaults to {start_ind}\n")
        tmp = re.sub("[^-\d]", "", tmp)
        if tmp == "":
            ind = start_ind
            if ind not in list(range(-len(hreflist), len(hreflist))):
                ind = -1  #dont terminate by default
        elif int(tmp) not in list(range(-len(hreflist), len(hreflist))):
            print("user terminated processing")
            break
        else:
            ind = int(tmp)  # process the specified one
            start_ind = ind  # use for next loop
        driver.execute_script(f"window.open('about:blank', 'current');")
        driver.switch_to.window("current")
        driver.get(hreflist[ind])
        driver.implicitly_wait(120)  # needs to be called per session
        if processtype == "C":
            processtypeC(driver)
        elif processtype == "S":
            processtypeS(driver)
        elif processtype == "R":
            processtypeR(driver)
        driver.switch_to.window(mainhandle)
        print("********** next to update hreflist **********")
        driver.get(pagehref)
        driver.implicitly_wait(120)  # needs to be called per session
        hreflist = get_hreflist(driver, processtype)


def processall(driver: webdriver):
    print("press C\t to process type 协同")
    print("press S\t to process type 发文")
    print("press R\t to process type 收文")
    processtype = input("choose process type...\n")
    if processtype not in todo_pageref_dict.keys():
        input("defaults to 收文 type R, press to confirm...\n")
        processtype = "R"
    processallbytype(driver, processtype)


if __name__ == "__main__":
    driver = initdriver.init()
    while True:
        processall(driver)
        loop = input("continue to process？ else type N or n...\n")
        loop = loop.lower()
        if loop == "n":
            break
    # TODO background watch todolist