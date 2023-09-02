#!/usr/bin/env python
# coding: utf-8
from datetime import datetime
import os
import pandas as pd


def str_len(s):
    row_l = len(s)
    utf8_l = len(s.encode('utf-8'))
    return (utf8_l - row_l) / 2 + row_l


def splitsheet(fn, resfn):
    df=pd.read_excel(fn,index_col=0)
    groupdf = df.groupby("sheet")

    sheetnames = list(groupdf.groups.keys())
    dflist = []

    for n in sheetnames:
        sheetdf = groupdf.get_group(n).copy()
        sheetdf.reset_index(inplace=True, drop=True)
        sheetdf.index += 1
        sheetdf.loc[:, "序号"] = sheetdf.index
        dflist = dflist + [sheetdf]

    Excelwriter = pd.ExcelWriter(resfn, engine="xlsxwriter")
    for i, sheetdf in enumerate(dflist):
        sheetdf.to_excel(Excelwriter, sheet_name=sheetnames[i], index=False)
        worksheet = Excelwriter.sheets[sheetnames[i]]  # pull worksheet object
        for idx, col in enumerate(sheetdf):  # loop through all columns
            series = sheetdf[col]
            max_len = max((
                series.astype(str).map(str_len).max(),  # len of largest item
                str_len(str(series.name))  # len of column name/header
            )) + 2  # adding a little extra space
            worksheet.set_column(idx, idx, max_len)  # set column width
    Excelwriter.save()


if __name__ == "__main__":
    rawexcelfn = r"C:\Users\Amy19\Documents\科务会材料&总结计划\例会资料收集\raw办公例会资料.xlsx"
    excelfn = r"C:\Users\Amy19\Documents\科务会材料&总结计划\例会资料收集\res办公例会资料.xlsx"  # sheets 最多有10个，包括5个总结、5个计划
    splitsheet(rawexcelfn, excelfn)