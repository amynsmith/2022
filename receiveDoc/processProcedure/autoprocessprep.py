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
from processProcedure.processall import get_hreflist
from processProcedure.processall import todo_pageref_dict, selector_btn_dict
from processProcedure.processall import id_iframe, selector_totalnum, selector_totalnum_coop, selector_todolist_checkbox

from prepareData.scrapeopinions import opinionsbytype
from utils.textpurify import extract_names_from_niban
from utils.constants import translatenames
from prepareData.evaluateopinions import add_nextnames
from prepareData.loaddf import simpleloaddf
from prepareData.malicious import write_nibansheet

import logging.config
import yaml
logger = logging.getLogger('verboseLogger')


def preparebytype(processtype, opinionfn, nibanfn, times_zoomout=3):
    pagehref = todo_pageref_dict[processtype]
    driver = initdriver.init()
    driver.get(pagehref)
    driver.implicitly_wait(120)  # needs to be called per session
    for n in range(times_zoomout):
        pyautogui.keyDown('ctrl')
        pyautogui.press('-')
        pyautogui.keyUp('ctrl')

    mainhandle = driver.current_window_handle
    hreflist = get_hreflist(driver, processtype)
    if len(hreflist)<1:
        logger.info("no left todos, exit")
        return
    # 按时间顺序处理，和之前页面抓取顺序相反

    reversed_hreflist = hreflist.copy()
    reversed_hreflist.reverse()
    if os.path.isfile(nibanfn):
        logger.info("based on original niban file")
        # full load
        tmpdf, nibandf = simpleloaddf(processtype, nibanfn, colfilter="toproc")
    else:
        nibandf = pd.DataFrame()
    if len(nibandf)<1:
        logger.info("no prev niban records")

    # initialize
    nibandf["toproc"]=False
    nibandf["opinionsheet"] = ""
    logger.info("start to record opinion from scratch...")
    if os.path.isfile(opinionfn):
        os.remove(opinionfn)  # record opinion from scratch

    for ind, href in enumerate(reversed_hreflist):
        sheetname = str(ind+1).zfill(3)
        logger.info(f"\ncurrent reversed href ind : {ind}, opinion sheetname: {sheetname}")
        logger.debug(href)
        driver.execute_script(f"window.open('about:blank', 'current');")
        driver.switch_to.window("current")
        driver.get(href)
        driver.implicitly_wait(120)  # needs to be called per session
        # fully rewrite, no more judge on rowind
        ########### per href
        output = opinionsbytype(driver, processtype, opinionfn, sheetname)
        if processtype == "R":
            title, href, recno, niban = output
            # get people name cols from niban string
            rawnames = extract_names_from_niban(niban)
            resnames = translatenames(rawnames)  # unknown,departnames,midnames,topnames
            # combine and write to excel
            rowdata = [title, href, recno] + [niban] + [",".join(n) for n in resnames]
            cols = ["title", "href", "recno", "niban", "unknown", "departnames", "midnames", "topnames"]
            # ValueError: If using all scalar values, you must pass an index
            tmpdf = pd.DataFrame([dict(zip(cols, rowdata))])
            tmpdf.set_index("recno", drop=False, inplace=True)
        else:
            # processtype == "S"
            title, href, person, org, starttime, isofficial, niban = output
            # get people name cols from niban string
            rawnames = extract_names_from_niban(niban)
            resnames = translatenames(rawnames)  # unknown,departnames,midnames,topnames
            # combine and write to excel
            rowdata = output + [",".join(n) for n in resnames]
            cols = ["title", "href", "person", "org", "starttime", "isofficial", "niban", "unknown", "departnames",
                    "midnames", "topnames"]
            tmpdf = pd.DataFrame([dict(zip(cols, rowdata))])
            tmpdf.set_index(["starttime", "person"], drop=False, inplace=True)
        # decide which one to keep
        if tmpdf.index.isin(nibandf.index.to_numpy()):
            existingniban = nibandf.loc[tmpdf.index].niban.values[0]
            if existingniban == niban:
                logger.debug(f"no change of niban, ind {tmpdf.index}")
                # ################continue
                # append below, delete the original row
                tmpdf = nibandf.loc[tmpdf.index].copy()
                # remember to update the href !!!!
                tmpdf.href = href
                tmpdf.title = title
                logger.info(f"use prev niban, meanwhile update href and title")
                # nibandf.drop(tmpdf.index, axis=0, inplace=True)
            else:
                logger.warning(f"niban altered, ind {tmpdf.index}")
                logger.info(f"prev niban: {existingniban}")
                logger.info(f"cur niban: {niban}")
                # nibandf.drop(tmpdf.index, axis=0, inplace=True)
            # always delete the original row
            nibandf.drop(tmpdf.index, axis=0, inplace=True)
        tmpdf["toproc"] = True  # in present remaining hreflist
        tmpdf["opinionsheet"] = sheetname
        nibandf = pd.concat([nibandf, tmpdf], ignore_index=False)

    # if not os.path.isfile(nibanfn):
    #     logger.debug("new niban file must be created")
    #     nibandf.to_excel(nibanfn, sheet_name=processtype, index=False)
    # else:
    #     logger.debug(f"niban file exists, now updating it")
    #     with pd.ExcelWriter(nibanfn, engine="openpyxl", mode="a", if_sheet_exists="replace") as Excelwriter:
    #         nibandf.to_excel(Excelwriter, sheet_name=processtype, index=False)
    write_nibansheet(processtype, nibandf, nibanfn)

    driver.quit()
    input("modify niban, edit column 'unknown' (selectively fill the cell to skip the row).\n"
          "Extreme attention to replacement of leaders. Plus addition of names. Then press to continue...\n")
    add_nextnames(processtype, opinionfn, nibanfn)
    return


if __name__ == "__main__":
    with open('../config.yaml', 'r') as f:
        config = yaml.safe_load(f.read())
        logging.config.dictConfig(config)
    logger = logging.getLogger('verboseLogger')

    resfolder = r"C:\Users\Amy19\PycharmProjects\receiveDoc\prepareData"
    opinionfn = os.path.join(resfolder, "records_opinion.xlsx")  # multi sheets corresponding to procedures
    nibanfn = os.path.join(resfolder,
                           "extract_niban.xlsx")  # single sheets with multi rows each corresponding to a procedure

    times_zoomout=3
    # processtype = "R"
    processtype = "S"
    preparebytype(processtype, opinionfn, nibanfn, times_zoomout)

    # modify niban based on opinions
    # input("modify and rename niban, then press to continue...\n")

    # then evaluate opinons and add nextnames (3 cols) to niban
    # modified_nibanfn = os.path.join(resfolder,"extract_niban_modified.xlsx")
    # add_nextnames(processtype, opinionfn, modified_nibanfn)
