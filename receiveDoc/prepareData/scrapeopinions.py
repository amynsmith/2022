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

from startProcedure.bulkreceivefile import selector_doc_title
from processProcedure.advancedoperation import getorgfromname

id_commentiframe = "componentDiv"
id_signiframe = "zwIframe" #inner  #for type R
selector_commentnum = "//span[contains(text(), '处理人意见区')]"
selector_comments = "*//div[starts-with(@id,'replay')]"
selector_brief_opinion = ".//div[./a]/span"
selector_detail_opinion = ".//ul[starts-with(@id, 'ulcomContent')]"
# for type R
selector_recno = "#field0011"
# for type S
selector_title = ".//input[@id='subject']"
selector_startperson = ".//input[@id='startMemberName']"
selector_starttime = ".//input[@id='createDate']"
# for layout header, to distinguish between formal and informal type of S
selector_headcells = ".//table[@class='xdLayout']/tbody/tr[1]/td"

import logging.config
import yaml
logger = logging.getLogger('verboseLogger')


def getrecno(driver):
    # 文单区域
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, id_commentiframe)))
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, id_signiframe)))
    # 收文标题、收文编号
    recno = driver.find_element(By.CSS_SELECTOR, selector_recno).get_property("innerText")
    driver.switch_to.default_content()  # don't mess with later selenium operation!!!
    return recno


def getnthopinion(driver, nth=-1):
    # 处理意见区
    driver.switch_to.default_content()
    driver.switch_to.frame(id_commentiframe)
    comment_elements = driver.find_elements(By.XPATH, selector_comments)
    commentnum = len(comment_elements)
    # 获取 brief_opinion
    # ['同意', '', '2022-08-07 09:13', '拟办']
    nth = int(nth)
    if nth not in list(range(-commentnum, commentnum)):
        logger.debug("nth of opinion defaults to -1")
        nth = -1
    comment = driver.find_elements(By.XPATH, selector_comments)[nth]
    tmpspans = comment.find_elements(By.XPATH, selector_brief_opinion)
    brief_opinion = [s.get_property("innerText").strip() for s in tmpspans]
    str_brief_opinion = " ".join(brief_opinion)
    driver.switch_to.default_content()  # don't mess with later selenium operation!!!
    return str_brief_opinion


# opinion excel consists of multiple sheets, each sheet relates to one href link
# def opinionsR(driver, opinionfn, sheetname):
#     href=driver.current_url
#     # 文单区域
#     WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, id_commentiframe)))
#     WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, id_signiframe)))
#     # 收文标题、收文编号
#     title = driver.find_element(By.CSS_SELECTOR, selector_doc_title).get_property("innerText")
#     recno = driver.find_element(By.CSS_SELECTOR, selector_recno).get_property("innerText")
#     # 处理意见区
#     driver.switch_to.default_content()
#     driver.switch_to.frame(id_commentiframe)
#     comment_elements = driver.find_elements(By.XPATH, selector_comments)
#     commentnum = len(comment_elements)
#     # 处理意见条数
#     # commentnumstr = driver.find_element(By.XPATH, selector_commentnum).get_property("innerText")
#     # commentnum =  int(re.sub(r"\D", "", commentnumstr))
#     # 记录所有处理意见
#     persons, brief_opinions, detail_opinions=[],[],[]
#     for i in range(commentnum):
#         # 获取 brief_opinion
#         # ['同意', '', '2022-08-07 09:13', '拟办']
#         comment = driver.find_elements(By.XPATH, selector_comments)[i]
#         persons.append(comment.find_element(By.TAG_NAME, "a").get_property("title"))
#         tmpspans = comment.find_elements(By.XPATH, selector_brief_opinion)
#         brief_opinions.append([s.get_property("innerText").strip() for s in tmpspans])
#         # 获取 detail_opinion
#         # eg1 呈X总阅；请X总阅；请XX阅处。
#         # eg2 已阅
#         # 若意见区有内容
#         if comment.get_property("childElementCount") >2:
#             tmpdetail = comment.find_element(By.XPATH, selector_detail_opinion).get_property("innerText")
#             tmpdetail = tmpdetail.strip("发自移动客户端")
#             tmpdetail=tmpdetail.replace("\n","")
#         # 文书节点一般没有意见区
#         else:
#             tmpdetail=""
#         detail_opinions.append(tmpdetail)
#     cols = ["person","brief_opinion","detail_opinion"]
#     opiniondf = pd.DataFrame(dict(zip(cols,[persons, brief_opinions, detail_opinions])))
#     opiniondf["str_brief_opinion"]= opiniondf.brief_opinion.apply(lambda x: " ".join(x))
#     # 取最新拟办意见
#     niban = opiniondf[opiniondf.str_brief_opinion.str.endswith("拟办")].detail_opinion.values[-1]
#     # 单条流程，记录意见区
#     # mode="a" 追加sheet
#     # if_sheet_exists，若给定 sheet 已存在，则覆写
#     # Append mode is not supported with xlsxwriter!
#     with pd.ExcelWriter(opinionfn, engine="openpyxl", mode="a", if_sheet_exists="replace") as Excelwriter:
#         opiniondf.to_excel(Excelwriter, sheet_name=sheetname, index=False)
#     # return processtype?
#     return title, href, recno, niban


# return niban columns:
# title	href
# for type R: recno
# for type S: person,org,starttime
# niban
# #### later cal ##### unknown	departnames	midnames	topnames	toproc	next_1	next_2	next_3	procdone
def opinionsbytype(driver, processtype, opinionfn, sheetname):
    href=driver.current_url
    output = []
    if processtype == "R":
        # 文单区域
        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, id_commentiframe)))
        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, id_signiframe)))
        # 收文标题、收文编号
        title = driver.find_element(By.CSS_SELECTOR, selector_doc_title).get_property("innerText")
        recno = driver.find_element(By.CSS_SELECTOR, selector_recno).get_property("innerText")
        output=[title, href, recno]
    else:
        # processtype == "S"
        # 文单区域
        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, id_commentiframe)))
        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, id_signiframe)))
        # TODO！！！！！！！！！！！！！！  loop through first row 标题，也适用于收文文单
        # 判断是否为红头发文
        # $x(".//table[@class='xdLayout']/tbody/tr[1]/td[last()]")[0].innerText
        # '（送审稿）'
        headcells = driver.find_elements(By.XPATH, selector_headcells)
        tmplist = [e.get_property("innerText").strip() for e in headcells]
        header = " ".join([s for s in tmplist if s != ""])
        isofficial = [True if header.endswith("（送审稿）") else False][0]
        # title,name,startorg,starttime
        href = driver.current_url
        driver.switch_to.default_content()
        title = driver.find_element(By.XPATH, selector_title).get_property("value")
        personstr = driver.find_element(By.XPATH, selector_startperson).get_property("value")
        person, org = getorgfromname(personstr)
        starttime = driver.find_element(By.XPATH, selector_starttime).get_property("value")
        output = [title, href, person,org,starttime, isofficial]
    # 处理意见区
    driver.switch_to.default_content()
    driver.switch_to.frame(id_commentiframe)
    comment_elements = driver.find_elements(By.XPATH, selector_comments)
    commentnum = len(comment_elements)
    # 处理意见条数
    # commentnumstr = driver.find_element(By.XPATH, selector_commentnum).get_property("innerText")
    # commentnum =  int(re.sub(r"\D", "", commentnumstr))
    # 记录所有处理意见
    persons, brief_opinions, detail_opinions=[],[],[]
    for i in range(commentnum):
        # 获取 brief_opinion
        # ['同意', '', '2022-08-07 09:13', '拟办']
        comment = driver.find_elements(By.XPATH, selector_comments)[i]
        persons.append(comment.find_element(By.TAG_NAME, "a").get_property("title"))
        tmpspans = comment.find_elements(By.XPATH, selector_brief_opinion)
        brief_opinions.append([s.get_property("innerText").strip() for s in tmpspans])
        # 获取 detail_opinion
        # eg1 呈X总阅；请X总阅；请XX阅处。
        # eg2 已阅
        # 若意见区有内容
        if comment.get_property("childElementCount") >2:
            tmpdetail = comment.find_element(By.XPATH, selector_detail_opinion).get_property("innerText")
            tmpdetail = tmpdetail.strip("发自移动客户端")
            tmpdetail=tmpdetail.replace("\n","")
        # 文书节点一般没有意见区
        else:
            tmpdetail=""
        detail_opinions.append(tmpdetail)
    cols = ["person","brief_opinion","detail_opinion"]
    opiniondf = pd.DataFrame(dict(zip(cols,[persons, brief_opinions, detail_opinions])))
    opiniondf["str_brief_opinion"]= opiniondf.brief_opinion.apply(lambda x: " ".join(x))
    # 取最新拟办意见
    tmpnibans = opiniondf[opiniondf.str_brief_opinion.str.endswith("拟办")].detail_opinion.values
    if len(tmpnibans) <1:
        niban=""
    else:
        niban = tmpnibans[-1]
    output.append(niban)
    # 单条流程，记录意见区
    # mode="a" 追加sheet
    # if_sheet_exists，若给定 sheet 已存在，则覆写
    # Append mode is not supported with xlsxwriter!
    if not os.path.isfile(opinionfn):
        logger.debug("need to create opinion file")
        opiniondf.to_excel(opinionfn, sheet_name=sheetname, index=False)
    else:
        logger.debug(f"in opinion file, current sheetname {sheetname}")
        with pd.ExcelWriter(opinionfn, engine="openpyxl", mode="a", if_sheet_exists="replace") as Excelwriter:
            opiniondf.to_excel(Excelwriter, sheet_name=sheetname, index=False)
    return output