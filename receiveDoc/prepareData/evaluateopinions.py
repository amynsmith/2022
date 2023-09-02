import pandas as pd
import datetime as dt
import numpy as np

import glob
import re
import os, sys

from prepareData.loaddf import simpleloaddf
import logging.config
import yaml
logger = logging.getLogger('verboseLogger')

# agree, disagree, notyet
def get_approvals(names, opiniondf):
    # for each name, find the last approval choice
    agree, disagree, notyet, other=[],[],[],[]
    opiniondf = opiniondf[~(opiniondf.brief_opinion.str.contains("拟办|文书管理", regex=True))]
    for n in names:
        if n not in opiniondf.person.values:
            notyet.append(n)
        else:
            tmpstr = opiniondf[opiniondf.person==n].str_brief_opinion.values[-1]
            tmpstr = tmpstr.split()[0]
            if tmpstr == "同意":
                agree.append(n)
            elif tmpstr == "不同意":
                disagree.append(n)
            else:
                print(f"unexpected attitude other than agree/disagree: {tmpstr}")
                other.append(n)
    return agree, disagree, notyet, other


def torollback(disagree, amongnames, opiniondf):
    # if len(disagree) <1:
    #     return False
    for n in disagree:
        startind = opiniondf[opiniondf.person==n].iloc[-1].name  #consider the last disagree opinion
        laterdf = opiniondf[opiniondf.index > startind]
        laternames = laterdf.person.values
        tmpset = set(laternames) - set(amongnames)
        # already rolled back, then forwarded to me
        # either the disagreed person or me operated rollback
        if len(tmpset)>0:
            return False
        # no yet rolled back, first sent to me
        else:
            return True

# eg:
# orderednames: [['XX', 'XX'], ['XX', 'XX', 'XX'], ['XX', 'XX']]
# nextnames: ['', 'YY', '']
# 实际此流程无需YY审批
# applies to processtype "S" only
def isspecialcase(orderednames, nextnames):
    # eg:
    # a1+a2,(b,c) -> a1,( , )
    # for i in range(2):  # i refers to former groups but the last one
    #         for j in range(i+1, 3):  # j refers to the latter group, including the last one
    #             print(i,j)
    # 按流转顺序，以层级组为单位
    for i in range(len(nextnames)-1):  # i refers to former groups but the last one
        # 此层级组成员全部处理完毕
        if len(nextnames[i]) == 0:
            continue
        # 此层级组未全数通过
        elif len(nextnames[i]) < len(orderednames[i]):
            # 后续层级组是否有处理，有则疑似流程节点跳跃
            for j in range(i+1, len(nextnames)):  # j refers to the latter group, including the last one
                if len(nextnames[j]) != len(orderednames[j]):
                    return True
    return False



# carried out OFFLINE
# then MUST check by hand
# on the basis of modification of nibandf
def add_nextnames(processtype, opinionfn, nibanfn):
    # nibandf = pd.read_excel(nibanfn)
    # nibandf.fillna("", inplace=True)
    loadednibandf, nibandf = simpleloaddf(processtype, nibanfn, colfilter="toproc")
    for row in loadednibandf.itertuples(index=True):
        ind = row.Index  # capitalized "I"
        opinionsheet = loadednibandf.loc[ind, "opinionsheet"]
        logger.info(f"\ncurrent niban at ind {ind}, corresponding opinionsheet {opinionsheet}")
        opiniondf = pd.read_excel(opinionfn, sheet_name=opinionsheet)
        # "unknown","departnames","midnames","topnames"
        tmpnames = [row.unknown, row.departnames, row.midnames, row.topnames]
        nameslist = [tmpstr.split(",") for tmpstr in tmpnames]
        unknown, departnames, midnames, topnames = nameslist
        if unknown != [""]:
            # skip
            logger.warning("cell 'unknown' not empty, skip")
            continue
        else:
            logger.info("unknown names cleared, adding nextnames")
            # 所有待处理人已确定
            if processtype == "R":
                orderednames = [topnames, midnames, departnames]
            else:
                # processtype == "S"
                orderednames = [departnames, midnames, topnames]
            nextnames = []
            for names in orderednames:
                agree, disagree, notyet, other = get_approvals(names, opiniondf)
                # selrange is leader/depart
                nextnames.append(list(set(names) - set(agree)))
                logger.info(f"agree: {agree}")
                logger.info(f"disagree: {disagree}")
                logger.info(f"notyet: {notyet}")
                logger.info(f"other: {other}")
                if (len(disagree)>0) and torollback(disagree, names, opiniondf):
                    logger.info("ATTENTION, rollback needed")
                    nextnames = ["","",""]
                    break
            # 对于发文类，是否有遗漏，或者人为跳过不需审批的情况
            if (processtype == "S") and isspecialcase(orderednames, nextnames):
                logger.info("ATTENTION, person omitted")
                nextnames = ["", "", ""]

            nextnamestrs = [",".join(n) for n in nextnames]
            nibandf.loc[ind, ["next_1", "next_2", "next_3"]] = nextnamestrs
            logger.info(f"nextnamestrs: {nextnamestrs}\n")
    # nibandf.to_excel(nibanfn, sheet_name=processtype, index=False)
    with pd.ExcelWriter(nibanfn, engine="openpyxl", mode="a", if_sheet_exists="replace") as Excelwriter:
        nibandf.to_excel(Excelwriter, sheet_name=processtype, index=False)


if __name__ == "__main__":
    resfolder = r"C:\Users\Amy19\PycharmProjects\receiveDoc\prepareData"
    opinionfn=os.path.join(resfolder,"records_opinion.xlsx")      # multi sheets corresponding to procedures
    nibanfn=os.path.join(resfolder,"extract_niban.xlsx")
    # modified_nibanfn = os.path.join(resfolder,"extract_niban_modified.xlsx")   # single sheets with multi rows each corresponding to a procedure
    # modify niban based on opinions
    # then evaluate opinons and add nextnames (3 cols) to niban
    processtype = "S"
    # processtype = "R"
    add_nextnames(processtype,opinionfn, nibanfn)