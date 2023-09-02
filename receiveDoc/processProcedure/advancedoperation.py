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
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException,UnexpectedAlertPresentException,ElementClickInterceptedException

import os, sys

sys.path.append(os.path.abspath(".."))
# import utils.initWebDriver as initdriver
from utils.constants import getallnames, getpolicy,getselrange
from utils.constants import people_depart

selector_outerframe = "iframe[id^='layui-layer-iframe']"  # also useful for selecting people
displaytablist = ["部门", "组"]
id_insertmodedict = {"a": "processMode_parallel", "b": "processMode_nextparallel"}
selector_additem = "//*[@id='selectPeopleTable']//span[@title='选择']"
selector_deleteitem = "//*[@id='selectPeopleTable']//span[@title='删除']"


# TODO 加签 人员 节点权限
# TODO 指定回退 流程重走 或 直接提交到我，协同类必填回退意见


orgtreelist = ["XX分公司",
               "XX分公司",
               "YY经理部",
               "ZZ经理部"]

orgabbrlist= ["XXXX",
              "XXXX",
              "XXYY",
              "XXZZ"]

# deprecated
sysgrouplist = [("领导班子及总助（收文加签范围，最新版）", "leader"), ("分公司总部部门负责人（正职）", "depart")]


# eg: XXX(YYYY)
def getorgfromname(name:str):
    for ind,abbr in enumerate(orgabbrlist):
        appendstr="("+ abbr +")"
        if appendstr in name:
            realname = name.rsplit(sep=appendstr,maxsplit=1)[0]
            return realname, orgtreelist[ind]
    appended = any( "".join(["(",abbr,")"]) in name for abbr in orgabbrlist)
    if not appended:
        return name, ""



# already is default content
def getprocneeds(driver: webdriver, processtype: str, selrange: str):
    if processtype == "S":
        stnameinfo = driver.find_element(By.ID, "panleStart").get_property("innerText")
        stcompany = re.search("(?<=\().*(?=\))", stnameinfo).group()
        if stcompany.endswith("XX"):
            templatecreator = "tj"
        elif stcompany.endswith("YY"):
            templatecreator = "sd"
        elif stcompany.endswith("ZZ"):
            templatecreator = "xa"
        else:
            templatecreator = "bf"
        # selrange = "midleader"
        policy = getpolicy(processtype, selrange, templatecreator)
    else:
        # processtype == "R":
        # print("press 0\t to apply policy of leaders")
        # print("press 1\t to apply policy of departments")
        # choice = input("type choice number...\n")
        # selrange = ["leader" if choice == "0" else "depart"][0]
        policy = getpolicy(processtype, selrange, templatecreator="")
    print("press A or a\t to choose 并发")
    print("press B or b\t to choose 与下一节点并发")
    choice = input("type choice number, defaults to b...\n")
    choice = choice.lower()
    # defaults to 与下一节点并发
    ind = ["b" if choice not in ["a", "b"] else choice][0]
    id_insertmode = id_insertmodedict[ind]
    return policy, id_insertmode


# def selectpeople_by_search(driver: webdriver):
#     names = getallnames(selrange="leader")
#     selrange = getselrange(names)
#     print(f"getselrange: {selrange}")
#     input_element = driver.find_element(By.ID, "q")
#     for n in names:
#         print(f"current name:\t {n}")
#         found = False
#         for org in orgtreelist:
#             org_xpath = "//*[@id='accountContentChild']//a[.='" + org + "']"
#             orgroot_xpath = "//*[@id='searchTable']//a[.='" + org + "']"
#             # click to change to org root
#             # collapse org tree
#             driver.find_element(By.ID, "select_input_div").click()
#             driver.find_element(By.XPATH, org_xpath).click()
#             driver.find_element(By.XPATH, orgroot_xpath).click()
#             # wait for orgroot to be selected
#             WebDriverWait(driver, 10).until(
#                 EC.text_to_be_present_in_element_attribute((By.XPATH, orgroot_xpath), "class", "selected"))
#             # start to search
#             input_element.clear()
#             input_element.send_keys(n)
#             input_element.send_keys(Keys.ENTER)
#             sel_element = driver.find_element(By.ID, "memberDataBody")
#             sel = Select(sel_element)
#             namesnum = sel_element.get_property("childElementCount")  # or length?
#             if namesnum < 1:
#                 print(f"not found in org {org}")
#             else:
#                 for opt in sel.options:
#                     opt_text = opt.get_property('innerText')
#                     name = opt_text.split()[0]
#                     if name == n:
#                         action = ActionChains(driver)
#                         action.double_click(opt).perform()
#                         print(f"selected from org {org}")
#                         found = True
#                         break
#                     else:
#                         print(f"search result {name} not match")
#                 if found:
#                     break
#     driver.switch_to.default_content()
#     driver.find_element(By.CSS_SELECTOR, "a[title='确定']").click()
#     return selrange

def selectpeople_by_search(driver: webdriver, nameslist:list):
    if len(nameslist)==0:
        # manually type in
        names = getallnames(selrange="leader")
        if len(names)==0:
            names = getallnames(selrange="depart")
        selrange = getselrange(names)
        print(f"getselrange: {selrange}")
    else:
        names=nameslist
        selrange=""
    input_element = driver.find_element(By.ID, "q")
    for n in names:
        print(f"current name:\t {n}")
        n, dstorg=getorgfromname(n)
        print(f"real name:\t {n} \ndst org:\t {dstorg}")
        found = False
        for org in orgtreelist:
            if (dstorg != "") and (dstorg != org):
                print(f"skip org {org}")
                continue
            org_xpath = "//*[@id='accountContentChild']//a[.='" + org + "']"
            orgroot_xpath = "//*[@id='searchTable']//a[.='" + org + "']"
            # click to change to org root
            # collapse org tree
            driver.find_element(By.ID, "select_input_div").click()
            driver.find_element(By.XPATH, org_xpath).click()
            driver.find_element(By.XPATH, orgroot_xpath).click()
            # wait for orgroot to be selected
            WebDriverWait(driver, 10).until(
                EC.text_to_be_present_in_element_attribute((By.XPATH, orgroot_xpath), "class", "selected"))
            # start to search
            input_element.clear()
            input_element.send_keys(n)
            input_element.send_keys(Keys.ENTER)
            sel_element = driver.find_element(By.ID, "memberDataBody")
            sel = Select(sel_element)
            namesnum = sel_element.get_property("childElementCount")  # or length?
            if namesnum < 1:
                print(f"not found in org {org}")
            else:
                for opt in sel.options:
                    opt_text = opt.get_property('innerText')
                    name = opt_text.split()[0]
                    if name == n:
                        action = ActionChains(driver)
                        action.double_click(opt).perform()
                        print(f"selected from org {org}")
                        found = True
                        break
                    else:
                        print(f"search result {name} not match")
                if found:
                    break
    if len(nameslist) == 0:
        driver.switch_to.default_content()
        driver.find_element(By.CSS_SELECTOR, "a[title='确定']").click()
    else:
        pass
    return selrange

# not in use
def selectpeople_by_sysgroup(driver: webdriver):
    for ind, l in enumerate(sysgrouplist):
        print(f"ind {ind}\t {l}")
    tmpstr = input("type ind for sysgroup, defaults to 0\n")
    ind = [0 if int(tmpstr) not in range(len(sysgrouplist)) else int(tmpstr)][0]
    sysgroup, selrange = sysgrouplist[ind]
    opt_xpath = "//*[@id='TeamDataBody']//option[@title='" + sysgroup + "']"
    opt = driver.find_element(By.XPATH, opt_xpath)
    opt_value = opt.get_property("value")
    sel = Select(driver.find_element(By.ID, "TeamDataBody"))
    sel.deselect_all()
    sel.select_by_value(opt_value)
    action = ActionChains(driver)
    action.double_click(opt).perform()
    driver.switch_to.default_content()
    driver.find_element(By.CSS_SELECTOR, "a[title='确定']").click()
    return selrange


# for restricted upload only
def selectpeople_by_branch(driver: webdriver, branchlist: list):
    for ind,org in enumerate(orgtreelist):
        cur_branch=branchlist[ind]
        print(f"current branch:\t {org}:{cur_branch}")
        if cur_branch is None:
            print("skip")
            continue
        else:
            org_xpath = "//*[@id='accountContentChild']//a[.='" + org + "']"
            orgroot_xpath = "//*[@id='searchTable']//a[.='" + org + "']"
            # click to change to org root
            # collapse org tree
            driver.find_element(By.ID, "select_input_div").click()
            driver.find_element(By.XPATH, org_xpath).click()
            driver.find_element(By.XPATH, orgroot_xpath).click()
            # wait for orgroot to be selected
            WebDriverWait(driver, 10).until(
                EC.text_to_be_present_in_element_attribute((By.XPATH, orgroot_xpath), "class", "selected"))
            # start to click
            cur_branches = cur_branch.split(" ")
            for br in cur_branches:
                if br == org:
                    driver.find_element(By.XPATH, selector_additem).click()
                    print(f"whole {org} selected")
                    break
                else:
                    # select departs below
                    br_xpath = "//*[@id='List1']//a[.='" + br + "']"
                    driver.find_element(By.XPATH, br_xpath).click()
                    driver.find_element(By.XPATH, selector_additem).click()
                    print(f"{br} selected")
    # driver.switch_to.default_content()
    # driver.find_element(By.CSS_SELECTOR, "a[title='确定']").click()   # beneath outerframe
    return


def selectnodeattr(driver: webdriver, processtype: str, groupflag: str, selrange: str, backtome: bool):
    policy, id_insertmode = getprocneeds(driver, processtype, selrange)
    # update frame id
    id_layoutframe = driver.find_element(By.CSS_SELECTOR, selector_outerframe).get_property("id")
    driver.switch_to.frame(id_layoutframe)
    # update policy
    opt_xpath = "//*[@id='policySelect']//option[contains(text(),'" + policy + "')]"
    opt_value = driver.find_element(By.XPATH, opt_xpath).get_property("value")
    sel = Select(driver.find_element(By.ID, "policySelect"))
    sel.select_by_value(opt_value)
    # click insertmode radio input
    driver.find_element(By.ID, id_insertmode).click()
    # ensure exmode to be all
    # only for selectpeople_by_sysgroup
    if groupflag == "部门":
        print("no choice of exemode when select people by search")
    elif groupflag == "组":
        id_exemode = "all_mode"
        driver.find_element(By.ID, id_exemode).click()
    # by default, back-to-me and back-to-start isn't checked, thus no need to double-check back-to-start
    if backtome is True:
        element = driver.find_element(By.ID, "backToMe")
        if element.get_property("checked") is False:
            element.click()
            print("backToMe now checked")
    driver.switch_to.default_content()
    driver.find_element(By.CSS_SELECTOR, "a[title='确定']").click()
    return


def advanceforward(driver: webdriver, processtype: str):
    backtome=True
    while True:
        # 点击加签
        driver.find_element(By.ID, "moreLabel").click()
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "a[title='加签']"))).click()
        id_layoutframe = driver.find_element(By.CSS_SELECTOR, selector_outerframe).get_property("id")
        driver.switch_to.frame(id_layoutframe)
        print("press 0\t to manually type and search names")
        print("press 1\t to add 分公司总部领导班子及总助")
        print("press 2\t to add 分公司总部部门负责人（正职）")
        tmpstr = input("type choice number, defaults to 0...\n")
        choice = ["0" if tmpstr not in ["0", "1", "2"] else tmpstr][0]

        # 弃用新oa自建“组”，只用“部门”页，按姓名、部门来选择
        # display = ["0" if tmpstr not in ["0", "1"] else tmpstr][0]
        # tabname = displaytablist[int(display)]
        groupflag = "部门"
        driver.find_element(By.XPATH, "//*[@id='tdPanel']//a[.='" + groupflag + "']").click()
        # if tabname == "部门":
        #     selrange = selectpeople_by_search(driver,[])
        # else:
        #     selrange = selectpeople_by_sysgroup(driver)
        if choice == "0":
            # manually type and search names
            selrange = selectpeople_by_search(driver, [])
            # groupflag = "部门"
        elif choice == "1":
            branchlist = ["领导班子 总经理助理", None, None, None]
            try:
                selectpeople_by_branch(driver, branchlist)
                # selenium.common.exceptions.UnexpectedAlertPresentException: Alert Text: "总经理助理"是否包含子部门?(确定表示包含，取消表示不包含)
                driver.switch_to.default_content()
                driver.find_element(By.CSS_SELECTOR, "a[title='确定']").click()
            except UnexpectedAlertPresentException:
                WebDriverWait(driver, 3).until(EC.alert_is_present())
                driver.switch_to.alert.accept()
                print("alert accepted")
                driver.switch_to.default_content()
                driver.find_element(By.CSS_SELECTOR, "a[title='确定']").click()
            selrange = "leader"
            groupflag="组"
        else:
            nameslist=[people_depart[k].value for k in people_depart.__members__.keys()]
            selectpeople_by_search(driver, nameslist)
            driver.switch_to.default_content()
            driver.find_element(By.CSS_SELECTOR, "a[title='确定']").click()
            selrange = "depart"
            # groupflag = "部门"

        # 选择加签节点属性
        try:
            WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "workflow_dialog_insertNodePage_id_main")))
        except UnexpectedAlertPresentException:
            WebDriverWait(driver, 3).until(EC.alert_is_present())
            driver.switch_to.alert.accept()
            print("alert accepted...")
            WebDriverWait(driver, 30).until(
                EC.visibility_of_element_located((By.ID, "workflow_dialog_insertNodePage_id_main")))

        selectnodeattr(driver, processtype, groupflag, selrange, backtome)
        backtome=False
        loop = input("continue to add nodes？ else type N or n...\n")
        loop = loop.lower()
        if loop == "n":
            break
    input("LAST, press to submit...\n")
    driver.find_element(By.CSS_SELECTOR, "input[value='提交']").click()