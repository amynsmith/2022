import pandas as pd
import numpy as np
import os
import glob
import re
# import datetime as dt
from datetime import datetime


# docmark isattach title recsource fpath pagenum pubdate recno recdate toreceive toupload

def gettypeind(recsource):
    if recsource == "X局":
        typeind = 0
    elif recsource == "X公司":
        typeind = 2
    elif recsource == "分公司":
        typeind = 3
    else:
        typeind = 1
    return typeind


# mode=toreceive or toupload or all
# def loaddf(fn, mode="toreceive"):
#     df = pd.concat(pd.read_excel(fn, sheet_name=None), ignore_index=True)
#     df = df.astype({"isattach": bool, "toreceive": bool, "toupload": bool}) ##############
#     # 对日期列应设置文本格式
#     # df.pubdate = df.pubdate.apply(lambda d: datetime.strptime(str(int(d)), "%Y%m%d"))
#     # df.recdate = df.recdate.apply(lambda d: datetime.strptime(str(int(d)), "%Y%m%d"))
#     df["pubdate"] = df["pubdate"].apply(lambda d: pd.to_datetime(d, format="%Y%m%d"))
#     df["recdate"] = df["recdate"].apply(lambda d: pd.to_datetime(d, format="%Y%m%d"))
#     df[["toreceive", "toupload"]] = df[["toreceive", "toupload"]].fillna(True)
#     df[["toreceive", "toupload"]] = df[["toreceive", "toupload"]].astype(bool)
#     colfilter = [mode if mode != "all" else ":"][0]
#     loadeddf = df[df[colfilter]]
#     return loadeddf


# mode=toreceive or toupload or all
def loaddf(fn, mode="toreceive"):
    df = pd.concat(pd.read_excel(fn, sheet_name=None), ignore_index=True)
    df = df.astype({"isattach": bool, "toreceive": bool, "toupload": bool})  ##############
    df[["toreceive", "toupload"]] = df[["toreceive", "toupload"]].fillna(True)
    # 对日期列应设置文本格式
    # df.pubdate = df.pubdate.apply(lambda d: datetime.strptime(str(int(d)), "%Y%m%d"))
    # df.recdate = df.recdate.apply(lambda d: datetime.strptime(str(int(d)), "%Y%m%d"))
    df.loc[:, "pubdate"] = pd.to_datetime(df.pubdate, format="%Y%m%d", errors="coerce")
    df.loc[:, "recdate"] = pd.to_datetime(df.recdate, format="%Y%m%d", errors="coerce")
    colfilter = [mode if mode != "all" else ":"][0]
    loadeddf = df[df[colfilter]]
    return loadeddf, df


# for auto process, load nibandf
def simpleloaddf(processtype, fn, colfilter="toproc"):
    # df = pd.concat(pd.read_excel(fn, sheet_name=None, dtype={"opinionsheet": str}), ignore_index=True)
    # ensure the sheet exists
    blankdf = pd.DataFrame()
    with pd.ExcelWriter(fn, engine="openpyxl", mode="a", if_sheet_exists="overlay") as Excelwriter:
        blankdf.to_excel(Excelwriter, sheet_name=processtype, index=False)
    # then fetch the sheet
    df = pd.read_excel(fn, sheet_name=processtype, dtype={"opinionsheet":str})
    if len(df)<1:
        # empty content in the sheet
        loadeddf = pd.DataFrame()
        df = pd.DataFrame()
        return loadeddf, df
    if processtype == "R":
        df.set_index("recno",drop=False,inplace=True)
    else:
        # processtype=="S"
        df.set_index(["starttime","person"],drop=False,inplace=True)
        df = df.astype({"isofficial":bool})
    df = df.astype({colfilter: bool})
    df[colfilter] = df[colfilter].fillna(True)
    df.fillna("", inplace=True)
    loadeddf = df[df[colfilter]]
    return loadeddf, df


# to receive or upload
def getinput(curdf, mode="toreceive"):
    bodyrow = curdf[~curdf.isattach]
    recsource = bodyrow.recsource.values[0]
    docmark = bodyrow.docmark.values[0]
    title = bodyrow.title.values[0]
    fpath = curdf[~curdf.isattach].fpath.values[0]
    attachpathlist = curdf[curdf.isattach].fpath.values
    if mode == "toreceive":
        recno = bodyrow.recno.values[0] + f"({datetime.today().year})"
        recdate = bodyrow.recdate.values[0]
        pagenum = bodyrow.pagenum.values[0]
        pubdate = bodyrow.pubdate.values[0]
        # inputlist = [recno, pd.Timestamp.strftime(recdate, "%Y-%m-%d"), recsource, docmark, str(pagenum),
        #              pd.Timestamp.strftime(pubdate, "%Y-%m-%d"), title]
        inputlist = [recno, pd.Timestamp(recdate).strftime("%Y-%m-%d"), recsource, docmark, str(int(pagenum)),
                     pd.Timestamp(pubdate).strftime("%Y-%m-%d"), title]
    else:
        # mode == "toupload"
        typeind = gettypeind(recsource)
        inputlist = [docmark + " " + title, typeind]
    return inputlist, fpath, attachpathlist


if __name__ == "__main__":
    # fn = r"C:\Users\Amy19\Desktop\收文20220402\respdflist_modified.xlsx"
    base_path = r"C:\Users\Amy19\Desktop\收文"
    basedir = base_path + pd.Timestamp.today().strftime("%Y%m%d")
    fn = basedir + "\\" + "respdflist_modified.xlsx"
    df = loaddf(fn, mode="toreceive")
    print(df.dtypes)
