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
import pyautogui
import os, sys


import utils.initWebDriver as initdriver
# from processProcedure.processall import get_hreflist
from processProcedure.processall import todo_pageref_dict, selector_btn_dict
# from processProcedure.processall import id_iframe, selector_totalnum, selector_totalnum_coop, selector_todolist_checkbox

from prepareData.scrapeopinions import opinionsbytype, getnthopinion, getrecno
# from utils.textpurify import extract_names_from_niban
# from utils.constants import translatenames
# from prepareData.evaluateopinions import add_nextnames
from utils.getcategory import getcategory
from processProcedure.processall import click_peoplebtn, click_archivebtn, click_submitbtn
from prepareData.loaddf import simpleloaddf

import logging.config
import yaml
logger = logging.getLogger('verboseLogger')


# already ensured that ALL people approved
def simpleadvance(driver, processtype="R"):
    if processtype=="R":
        lastopinion = getnthopinion(driver, nth=-1)
        if lastopinion.endswith("文书管理"):
            choice = "2"  #choose archive folder
            recno =getrecno(driver)
            category = getcategory(recno)
            click_archivebtn(driver, category, processtype)
        else:
            choice = "90"  #submit forward
            click_submitbtn(driver)
    else:
        # processtype == "S":
        lastopinion = getnthopinion(driver, nth=-1)
        requiredopinion = getnthopinion(driver, nth=-2)
        if lastopinion.endswith("成文") and requiredopinion.endswith("登记编号"):
            choice = "5"  # choose archive folder
            category=""
            click_archivebtn(driver, category, processtype)
        else:
            confirm = input("double check before submit, press S/s to skip...\n")
            if confirm.strip().lower() == "s":
                choice = "999"
                # driver.close()
            else:
                choice = "90"  #submit forward
                click_submitbtn(driver)
    return choice


# window main
def performbytype(processtype, nibanfn, times_zoomout=3):
    loadeddf, df = simpleloaddf(processtype, nibanfn, colfilter="toproc")
    df.loc[loadeddf.index, "procdone"]=False  # if not exists, add new column
    # filter out unclear niban records
    loadeddf = loadeddf[(loadeddf.unknown == "")]  # ensure
    if len(loadeddf)<1:
        logger.info("no todo record, exit")
        return

    pagehref = todo_pageref_dict[processtype]
    driver = initdriver.init()
    driver.get(pagehref)
    for n in range(times_zoomout):
        pyautogui.keyDown('ctrl')
        pyautogui.press('-')
        pyautogui.keyUp('ctrl')
    driver.implicitly_wait(120)  # needs to be called per session

    for row in loadeddf.itertuples(index=True):
        ind = row.Index  # capitalized "I"
        href = row.href
        logger.info(f"ind {ind}, totalnum {len(loadeddf)}")

        mainhandle = driver.current_window_handle
        driver.execute_script(f"window.open('about:blank', 'current');")
        driver.switch_to.window("current")
        driver.get(href)
        driver.implicitly_wait(120)  # needs to be called per session
        # ordered ahead

        nextnamestrs = loadeddf.loc[ind, ["next_1", "next_2", "next_3"]].values
        if processtype == "R":
            # orderednames = [topnames, midnames, departnames]
            if set(nextnamestrs) == {""}:
                # choice=="2" or "90"
                # dont click people btn
                choice = simpleadvance(driver, processtype)
                logger.info(f"applied choice: {choice}")
            else:
                # click people btn
                if (nextnamestrs[0]!="") or (nextnamestrs[1]!=""):
                    choice="0"  #leaders
                    # 合并前两列姓名，也即主职领导、分管领导节点，在收文流程中不需区分
                    tmpnames = nextnamestrs[0].split(",") + nextnamestrs[1].split(",")
                    names = [n for n in tmpnames if n!=""]
                    click_peoplebtn(driver, choice, names, processtype)
                else:
                    # nextnamestrs[2] != ""
                    names = nextnamestrs[2].split(",")
                    if "PAN" in names:
                        # direct submit
                        choice = simpleadvance(driver, processtype)
                        logger.info(f"applied choice: {choice}")
                    else:
                        choice = "1"  # departments
                        click_peoplebtn(driver, choice, names, processtype)
        else:
            # processtype == "S"
            # orderednames = [departnames, midnames, topnames]
            isofficial = loadeddf.loc[ind,"isofficial"]
            if set(nextnamestrs) == {""}:
                # dont click people btn
                # choice=="5" or "90"
                choice = simpleadvance(driver, processtype)
                logger.info(f"applied choice: {choice}")
            else:
                # click people btn
                if nextnamestrs[0] != "":
                    choice = ["3" if isofficial else "0"][0]  # depart
                    names = nextnamestrs[0].split(",")
                else:
                    if isofficial:
                        choice = "4"  # leader
                        # 合并后两列姓名，也即主职领导、分管领导节点，在套红流程中不需区分
                        tmpnames = nextnamestrs[1].split(",") + nextnamestrs[2].split(",")
                        names = [n for n in tmpnames if n != ""]
                    else:
                        if nextnamestrs[1]!="":
                            choice = "1"  # mid
                            names = nextnamestrs[1].split(",")
                        else:
                            # nextnamestrs[2]!=""
                            choice = "2"  # top
                            names = nextnamestrs[2].split(",")
                click_peoplebtn(driver, choice, names, processtype)

        WebDriverWait(driver, 30).until(EC.number_of_windows_to_be(1))
        # then switch back, otherwise next loop will run into exception:
        # NoSuchWindowException: Message: no such window: target window already closed
        driver.switch_to.window(mainhandle)
        df.loc[ind,"procdone"]=True
        df.loc[ind,"toproc"]=False
        # df.to_excel(nibanfn, index=False)
        with pd.ExcelWriter(nibanfn, engine="openpyxl", mode="a", if_sheet_exists="replace") as Excelwriter:
            df.to_excel(Excelwriter, sheet_name=processtype, index=False)
    return




if __name__ == "__main__":
    with open('../config.yaml', 'r') as f:
        config = yaml.safe_load(f.read())
        logging.config.dictConfig(config)
    logger = logging.getLogger('verboseLogger')

    resfolder = r"C:\Users\Amy19\PycharmProjects\receiveDoc\prepareData"
    # modified_nibanfn = os.path.join(resfolder, "extract_niban_modified.xlsx")
    nibanfn = os.path.join(resfolder, "extract_niban.xlsx")

    processtype = "R"
    # processtype = "S"
    performbytype(processtype, nibanfn)
