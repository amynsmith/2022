# import pandas as pd
# import datetime as dt
# import numpy as np
#
# import glob
# import re
# from selenium import webdriver
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from selenium.webdriver.chrome.options import Options
import os, sys

# sys.path.append(os.path.abspath(".."))
# import utils.initWebDriver as initdriver
# from processall import get_hreflist
# from processall import todo_pageref_dict, selector_btn_dict
# from processall import id_iframe, selector_totalnum, selector_totalnum_coop, selector_todolist_checkbox
# # for archive
# from processall import id_archivecheck, selector_outerframe, id_innerframe
# from processall import selector_docarchive, selector_receivearchive, selector_category_dict, selector_sendarchive


# print("press 0\t to choose leaders")
# print("press 1\t to choose departments")

# print("press 2\t to choose archive folder")
# print("press 90\t to submit forward")
# print("press 99\t to advance forward")

# modify niban ahead, to include more people according to additional opinion
# ensure niban names list fully covers requirements
# def inferchoiceR(driver: webdriver):
    # leaders (mid + top)
    # if the last comment is niban

    # or not all leaders approved, remember to check records of niban
    # get remaining leaders names

    # departs
    # after all leaders approved

    # submit forward
    # all leaders and departs both approved
    # last node is "承办"

    # choose archive folder
    # according to recno
    # brief opinion endswith "文书管理" but no "日志"


    # advance forward
    # print("press 0\t to manually type and search names")
    # print("press 1\t to add 分公司总部领导班子及总助")
    # print("press 2\t to add 分公司总部部门负责人（正职）")

from processProcedure.autoprocessprep import preparebytype
from processProcedure.autoprocessperform import performbytype

import logging.config
import yaml
logger = logging.getLogger('verboseLogger')


def autoprocessbytype(processtype: str,  opinionfn: str, nibanfn: str, toprep=True, times_zoomout=3):
    if toprep:
        preparebytype(processtype, opinionfn, nibanfn, times_zoomout)
    else:
        logger.info("prep skipped, based on previous output")
    performbytype(processtype, nibanfn, times_zoomout)


# based on previous niban file, may choose to recreate from scratch
# record opinion from scratch

if __name__ == "__main__":
    with open('../config.yaml', 'r') as f:
        config = yaml.safe_load(f.read())
        logging.config.dictConfig(config)
    logger = logging.getLogger('verboseLogger')

    resfolder = r"C:\Users\Amy19\PycharmProjects\receiveDoc\prepareData"
    opinionfn=os.path.join(resfolder,"records_opinion.xlsx")  # multi sheets corresponding to procedures
    nibanfn=os.path.join(resfolder,"extract_niban.xlsx")   # single sheets with multi rows each corresponding to a procedure

    # processtype = "R"
    processtype = "S"
    autoprocessbytype(processtype, opinionfn, nibanfn, toprep=True)
    # autoprocessbytype(processtype, opinionfn, nibanfn, toprep=False)

# def autoprocessallbytype(driver: webdriver, processtype: str):
#     pagehref = todo_pageref_dict[processtype]
#     driver.get(pagehref)
#     driver.implicitly_wait(120)  # needs to be called per session
#     mainhandle = driver.current_window_handle
#     hreflist = get_hreflist(driver, processtype)
#     start_ind = -1  # defaults to process in time order
#     while len(hreflist) > 0:
#         # print(f"ind number [out of range] leads to terminating this processtype {processtype}")
#         # tmp = input(f"type ind number\t to proceed, defaults to {start_ind}\n")
#         # tmp = re.sub("[^-\d]", "", tmp)
#         # if tmp == "":
#         #     ind = start_ind
#         #     if ind not in list(range(-len(hreflist), len(hreflist))):
#         #         ind = -1  #dont terminate by default
#         # elif int(tmp) not in list(range(-len(hreflist), len(hreflist))):
#         #     print("user terminated processing")
#         #     break
#         # else:
#         #     ind = int(tmp)  # process the specified one
#         #     start_ind = ind  # use for next loop
#
#         # 按时间顺序依次处理，不再手动指定
#         ind = start_ind
#         driver.execute_script(f"window.open('about:blank', 'current');")
#         driver.switch_to.window("current")
#         driver.get(hreflist[ind])
#         driver.implicitly_wait(120)  # needs to be called per session
#         if processtype == "C":
#             # autoprocesstypeC(driver)
#             pass
#         elif processtype == "S":
#             # autoprocesstypeS(driver)
#             pass
#         elif processtype == "R":
#             autoprocesstypeR(driver)
#         driver.switch_to.window(mainhandle)
#         print("********** next to update hreflist **********")
#         driver.get(pagehref)
#         driver.implicitly_wait(120)  # needs to be called per session
#         prevhref=hreflist[ind]
#         hreflist = get_hreflist(driver, processtype)
#         if hreflist[ind] == prevhref:
#             print("update hreflist twice...")
#             hreflist = get_hreflist(driver, processtype)




# def autoprocessall(driver: webdriver):
#     print("press C\t to process type 协同")
#     print("press S\t to process type 发文")
#     print("press R\t to process type 收文")
#     processtype = input("choose process type...\n")
#     if processtype not in todo_pageref_dict.keys():
#         input("defaults to 收文 type R, press to confirm...\n")
#         processtype = "R"
#         autoprocessallbytype(driver, processtype)