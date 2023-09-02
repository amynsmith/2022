#!/usr/bin/env python
# coding: utf-8

import pandas as pd
from docx import Document
from docxtpl import DocxTemplate

clslist = ["一、基础工作（5%）",
           "二、重点工作（40%）",
           "三、综合治理工作（20%）",
           "四、六个专项行动（15%）",
           "五、三年能力提升及成本压降（15%）",
           "六、其他工作和需各部门协调解决的相关事宜（5%）"]

varlist = ["总结：",
           "计划："]


def pastetoword(wordtemplatefn, fn, resfn):
    df = pd.read_excel(fn, sheet_name=None)
    for k in df.keys():
        df[k].fillna("", inplace=True)
    # 先总结后计划
    subdfnames = [v + c for v in varlist for c in clslist[:-1]]
    # 前面10张表格对应类别1-5
    # 最后1张表格对应类别6 不用表格填报，手动补填
    placeholders = ["section" + str(i) + tag for tag in ["a", "b"] for i in range(1, 6)]
    content = dict()
    for i in range(10):
        cursheet = subdfnames[i]
        if cursheet in df.keys():
            print(f"sheet exists: {cursheet}")
            if i < 5:
                # task.no task.desc task.prog
                tmpdf = df[cursheet][["序号", "重点事项", "完成情况"]]
                tmpdictlist = tmpdf.to_dict('records')
                tmptblcontent = [dict(zip(["no", "desc", "prog"], list(row.values()))) for row in tmpdictlist]
            else:
                # task.no task.desc task.method task.person task.t
                tmpdf = df[cursheet][["序号", "重点事项", "具体措施", "责任人", "时限"]]
                tmpdictlist = tmpdf.to_dict('records')
                tmptblcontent = [dict(zip(["no", "desc", "method", "person", "t"], list(row.values()))) for row in
                                 tmpdictlist]
        else:
            print(f"sheet not found: {cursheet}")
            tmptblcontent = [dict()]
        content[placeholders[i]] = tmptblcontent
    tpl = DocxTemplate(wordtemplatefn)
    tpl.render(content)
    tpl.save(resfn)



if __name__ == "__main__":
    wordtemplatefn = r"C:\Users\Amy19\Documents\科务会材料&总结计划\例会资料收集\template办公例会资料.docx"
    excelfn=r"C:\Users\Amy19\Documents\科务会材料&总结计划\例会资料收集\res办公例会资料.xlsx"
    wordfn=r"C:\Users\Amy19\Documents\科务会材料&总结计划\例会资料收集\res办公例会资料.docx"
    pastetoword(wordtemplatefn,excelfn,wordfn)