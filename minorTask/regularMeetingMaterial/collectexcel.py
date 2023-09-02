#!/usr/bin/env python
# coding: utf-8

import pandas as pd
from docx import Document
import os
import re
import glob as glob

# clslist=["一、基础工作（5%）",
# "二、重点工作（40%）",
# "三、综合治理工作（20%）",
# "四、六个专项行动（15%）",
# "五、三年能力提升及成本压降（15%）",
# "六、其他工作和需各部门协调解决的相关事宜（5%）"]
# varlist=["总结：",
# "计划："]
# # word中排列按（总结、计划）
# tblnamelist = [v+c for c in clslist[:-1] for v in varlist]
# # collect_subdfnames = collect_subdfnames+["六、其他工作和需各部门协调解决的相关事宜（5%）"]

# 最新要求改为：总结、计划表各一张，表内自行填写版块列
tblnamelist = ["总结", "计划"]

def extracttables(wordfn, tblnamelist):
    word=Document(wordfn)
    tbls = word.tables
    resdf=pd.DataFrame()
    if len(tbls) != len(tblnamelist):
        print("tables num not match!")
        return None
    for ind in range(len(tbls)):
        print(f"\ncurrent table: {tblnamelist[ind]}")
        curtb = tbls[ind]
        # 针对有多余空列的情况
        # 表头取列名，忽略空列，即列名必须非空
        colnames = [cell.text for cell in curtb.row_cells(0) if cell.text != ""]
        print(f"colnames:\t{colnames}")
        # 剔除空行和首行
        rownum = len(curtb.rows)
        skiprows=[0]  # 首先剔除首行
        for i in range(rownum):
            # 第三列重点事项必须非空
            tmpcell = curtb.column_cells(2)[i]
            if tmpcell.text == "":
                skiprows.append(i)
        if set(skiprows) == set(range(rownum)):
            print("no content in this table")
        else:
            print(f"skiprows: {skiprows}")
        userows = list(set(range(rownum)) - set(skiprows))
        userows.sort()
        # 填充当前表格数据
        curdata=dict.fromkeys(colnames, [])
        for i in range(len(colnames)):
            # curdata[colnames[i]]=[cell.text for cell in curtb.column_cells(i)][1:]
            allvalues = [cell.text for cell in curtb.column_cells(i)]
            curdata[colnames[i]] = [allvalues[ind] for ind in userows]

        # 添加sheet列
        # lendata = len(curtb.rows)-1
        lendata = len(userows)
        curdata["sheet"] = lendata* [tblnamelist[ind]]
        print("curdata extracted:")
        print(curdata)
        # if ind==0:
        #     resdf=pd.DataFrame(curdata)
        # else:
        #     resdf = pd.concat([resdf, pd.DataFrame(curdata)])
        resdf = pd.concat([resdf, pd.DataFrame(curdata)])
    return resdf


def collectexcel(wordfnlist,rawexcelfn):
    resdf = pd.DataFrame()
    for ind, wordfn in enumerate(wordfnlist):
        tmpperson = os.path.splitext(os.path.basename(wordfn))[0]
        person = re.sub("\W|\d", "", tmpperson)
        print("=====================\n"+person)
        tmpdf = extracttables(wordfn, tblnamelist)
        tmpdf["person"] = person
        print(len(tmpdf))
        # if ind == 0:
        #     resdf = tmpdf
        # else:
        #     resdf = pd.concat([resdf, tmpdf])
        resdf = pd.concat([resdf, tmpdf])
    resdf.reset_index(drop=True, inplace=True)
    # 删除多余空行
    # deleteind = resdf[resdf["重点事项"] == ""].index
    # resdf.drop(index=deleteind, axis=0, inplace=True)
    # sheet列移到最后
    resdf = resdf[[col for col in resdf if col not in ["sheet"]] + ["sheet"]]
    resdf.to_excel(rawexcelfn)



if __name__ == "__main__":
    rootfolder = r"C:\Users\Amy19\Documents\科务会材料&总结计划\例会资料收集\办公例会资料0830"
    tmpl = list(glob.iglob(rootfolder + "\\*.docx"))
    wordfnlist = [fpath for fpath in tmpl if not os.path.basename(fpath).startswith("办公例会资料")]
    rawexcelfn = r"C:\Users\Amy19\Documents\科务会材料&总结计划\例会资料收集\raw办公例会资料.xlsx"

    collectexcel(wordfnlist, rawexcelfn)