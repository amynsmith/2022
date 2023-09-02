#
# # TODO rename folder and filename, ignore "_resave.pdf"
# # TODO classify by pub department and sort from given num
# # TODO output df including attachlist

# TODO URGENT: iterate dirs and add column "attachlist"

import logging.config
import yaml
logger = logging.getLogger('verboseLogger')

import shutil

import pandas as pd
import numpy as np
import os
import re
import glob
import os, sys
from datetime import datetime

sys.path.append(os.path.abspath(".."))
from utils.extractfileinfo import extract_fileinfo, getsheet4receive
from utils.extractfileinfo import exclude_ptn, title_ptn, docmark_ptn
from utils.formatsheet import highlight_respdflist


# # already filtered with exclude_ptn
# def basicinfo(fpath):
#     dirname = os.path.dirname(fpath)
#     foldername = os.path.split(dirname)[-1]
#     fn = os.path.basename(fpath)
#     fname = os.path.splitext(fn)[0]
#
#     isattach = [True if re.search(docmark_ptn, fname) is None else False][0]
#     if isattach is True:
#         docmark = re.search(docmark_ptn, foldername).group()
#         title = fname.strip()
#     else:
#         docmark = re.search(docmark_ptn, fname).group()
#         title = re.search(title_ptn, fname).group().strip()
#     recsource = identifyissuer(docmark)
#     return docmark, isattach, title, recsource, fpath
#
#
# def generatedf(basedir, exclude_ptn):
#     rootpath = basedir + r"\**\*.*"
#     dfdata = []
#     for i in glob.iglob(rootpath, recursive=True):
#         if exclude_ptn.search(os.path.basename(i)) is not None:
#             # print(f"exclude {os.path.basename(i)}")
#             continue
#         # print(f"get basic info from {i}")
#         dfdata = dfdata + [list(basicinfo(i))]
#     df = pd.DataFrame(data=dfdata, columns=["docmark", "isattach", "title", "recsource", "fpath"])
#     return df
#
#
# def populatedf(basedir, resfn):
#     df = generatedf(basedir, exclude_ptn)
#     bodyfilter_ind = df[~df.isattach].index
#     # df.loc[bodyfilter_ind,:][df.loc[bodyfilter_ind,"docmark"].str.endswith("期")]
#     tmpdf = df.loc[bodyfilter_ind, :]
#     main_indlist = tmpdf[tmpdf.docmark.str.endswith("期")].index
#     part_indlist = tmpdf[~tmpdf.docmark.str.endswith("期")].index
#     resdf = populate_pubdatedf(df, df_indslices=[main_indlist, part_indlist], page_indlist=[0, -1])
#     resdf["recno"] = ""
#     resdf["recdate"] = "{d.year}{d.month}{d.day}".format(d=datetime.now())
#     resdf[["toreceive", "toupload"]] = ""
#     resdf.to_excel(resfn)
#     return resdf


# per file manipulation
# title comes from pdf filename, while folder name comes from download page
# no modification on folder name
def populatedf(basedir, resfn):
    rootpath = basedir + r"\**\*.*"
    dfdata = []
    for i in glob.iglob(rootpath, recursive=True):
        if (os.path.isfile(i) is False) or (exclude_ptn.search(os.path.basename(i)) is not None):
            logger.debug(f"exclude {os.path.basename(i)}")
            continue
        logger.info(f"===================\nextract info from {i}")
        dfdata = dfdata + [list(extract_fileinfo(i))]
    # docmark, isattach, title, recsource, fpath, pagenum, pubdate
    resdf = pd.DataFrame(data=dfdata, columns=["docmark", "isattach", "title", "recsource", "fpath", "pagenum", "pubdate"])
    resdf["dirname"] = resdf.fpath.apply(lambda f: os.path.dirname(f))
    resdf["recno"] = ""  # need double check
    resdf["recdate"] = "{d.year}{d.month}{d.day}".format(d=datetime.now())
    resdf[["toreceive", "toupload"]] = False  # need double check
    resdf.to_excel(resfn)
    return resdf


# docmark,isattach,recsource,pubdate are all checked beforehand
# toreceive,toupload are manually added
# modify resXXX.xlsx only, no file manipulation
def fill_docmark_forattach(fn):
    df = pd.read_excel(fn)
    df=df.astype({"isattach":bool, "toreceive":bool, "toupload": bool})
    # df["dirname"] = df.fpath.apply(lambda f: os.path.dirname(f))
    docmarkset = set(df[df.docmark.notna()].docmark)
    for docmark in docmarkset:
        dirname = df[df.docmark == docmark].dirname.values[0]
        ind = df[df.dirname == dirname].index
        df.loc[ind, "docmark"] = docmark
    df.to_excel(fn, index=False)
    return df


def rename_folder(dirname, docmark):
    foldername = os.path.split(dirname)[-1]
    basedir = os.path.dirname(dirname)
    if docmark in foldername:
        return dirname
    else:
        newfoldername = docmark+" "+foldername
        newdirname = os.path.join(basedir, newfoldername)
        shutil.move(dirname, newdirname)
        logger.debug(f"rename {foldername} to {newfoldername}")
        return newdirname


# docmark later filled in
# then rename folders and update fpath
# fn=basedir + "\\" + "respdflist.xlsx"
# updating fpath and modify fn
# start modification on folder name

# def update_fpath(basedir, fn):
#     df=pd.read_excel(fn)
#     docmarkset = set(df.docmark)
#     for docmark in docmarkset:
#         print(f"docmark {docmark}")
#         ori_fpathlist = df[df.docmark == docmark].fpath.values
#         dirname = df[df.docmark == docmark].dirname.values[0]
#         foldername = os.path.split(dirname)[-1]
#         ind = df[df.docmark == docmark].index
#         if docmark not in foldername:
#             newfoldername = docmark +" "+foldername
#             newdir = os.path.join(basedir, newfoldername)
#             print(f"need to rename folder {foldername} to {newfoldername}")
#             shutil.move(dirname, newdir)
#             new_fpathlist = []
#             for fpath in ori_fpathlist:
#                 fbase = os.path.basename(fpath)
#                 newfpath = os.path.join(basedir, newfoldername, fbase)
#                 new_fpathlist = new_fpathlist + [newfpath]
#             df.loc[ind, "fpath"] = new_fpathlist
#         else:
#             print(f"no need to rename folder {foldername}")
#     # update dirname
#     df.dirname = df.fpath.apply(lambda f: os.path.dirname(f))
#     df.to_excel(fn)
#     return df


def update_fpath(basedir, fn):
    rootpath = basedir + "\\*\\"
    tmp = list(glob.iglob(rootpath))
    folderlist = [os.path.dirname(i) for i in tmp]

    df = pd.read_excel(fn)
    df = df.astype({"isattach": bool, "toreceive": bool, "toupload": bool})

    bodydf = df[~df.isattach]
    # TODO SettingWithCopyWarning
    # bodydf["filename"] = bodydf[["docmark", "title"]].apply(lambda row: row.docmark + " " + row.title + ".pdf", axis=1)
    # bodydf.loc[:, "filename"] = bodydf[["docmark", "title"]].apply(lambda row: row.docmark + " " + row.title + ".pdf", axis=1)
    df.loc[bodydf.index, "filename"] = bodydf.apply(lambda row: row.docmark + " " + row.title + ".pdf", axis=1)
    bodydf = df[~df.isattach]  # NOT duplicate, because df is modified

    for ind, dirname in enumerate(folderlist):
        tmp = list(glob.iglob(dirname + "\\*.*"))
        # loop through files inside dir
        for f in tmp:
            # matches body file
            if os.path.basename(f) in bodydf.filename.values:
                foldername = os.path.split(dirname)[-1]
                filename = os.path.basename(f)
                tmpind = bodydf[bodydf.filename == filename].index
                docmark = df.loc[tmpind, "docmark"].values[0]
                newdirname = rename_folder(dirname, docmark)
                df.loc[tmpind, "dirname"] = newdirname
                continue
    # now update fpath for attach
    docmarkset = set(df.docmark)
    for docmark in docmarkset:
        logger.debug(f"docmark {docmark}")
        ind = df[df.docmark == docmark].index
        ori_fpathlist = df[df.docmark == docmark].fpath.values
        dirname = df[(df.docmark == docmark) & (~df.isattach)].dirname.values[0]
        # no file level rename
        new_fpathlist = []
        for fpath in ori_fpathlist:
            fbase = os.path.basename(fpath)
            newfpath = os.path.join(dirname, fbase)
            new_fpathlist = new_fpathlist + [newfpath]
        df.loc[ind, "fpath"] = new_fpathlist
        df.loc[ind, "dirname"] = dirname
    df.to_excel(fn, index=False)
    return df



# the last step
def fill_recno(fn, recfn):
    df = pd.read_excel(fn)
    df = df.astype({"isattach": bool, "toreceive": bool, "toupload": bool})

    recdf = pd.read_excel(recfn, sheet_name=None)
    keys = [k for k in recdf.keys() if k != "汇总"]
    start_recno_list = []
    for k in keys:
        sheetdf = recdf[k]
        if str(sheetdf.iloc[0]["文件标题"]) == 'nan':
            # no previous record
            start_recno = sheetdf.iloc[0, 0]
        else:
            # exists previous record
            sheetdf.dropna(axis=0, subset="文件标题", inplace=True)
            prev_recno = sheetdf.iloc[-1, 0]
            tmp = prev_recno.split("-")
            start_recno = tmp[0] + "-" + str(int(tmp[1]) + 1)
        start_recno_list = start_recno_list + [start_recno]
    start_recno_dict = dict(zip(keys, start_recno_list))
    df["sheet"] = df[["docmark", "recsource"]].apply(lambda row: getsheet4receive(row.docmark, row.recsource), axis=1)
    resdf = df.sort_values(by=["recsource", "sheet", "docmark", "isattach"])
    sheetset = set(resdf.sheet.values)
    for s in sheetset:
        logger.debug(s)
        start_recno = start_recno_dict[s]
        start_recno_ordernum = int(start_recno.split("-")[-1])
        sheetdf = resdf[resdf.sheet == s]
        bodydf = sheetdf[~sheetdf.isattach]
        tmpl = list(range(len(bodydf)))
        bodyrecnolist = [start_recno.split("-")[0] + "-" + str(tmp + start_recno_ordernum) for tmp in tmpl]
        for i in tmpl:
            logger.debug(bodydf.iloc[i].docmark)
            ori_index = bodydf.iloc[i].name
            resdf.loc[ori_index, "recno"] = bodyrecnolist[i]
    resdf.to_excel(fn, index=False)  # override
    return resdf




if __name__ == "__main__":
    with open('../config.yaml', 'r') as f:
        config = yaml.safe_load(f.read())
        logging.config.dictConfig(config)
    logger = logging.getLogger('verboseLogger')
    logging.getLogger("pdfminer").setLevel(logging.WARNING)

    # rootpath = r"C:\Users\Amy19\Desktop\收文20220402\**\*.pdf"
    # basedir = r"C:\Users\Amy19\Desktop\收文20220402"
    # resfn = r"C:\Users\Amy19\Desktop\收文20220402\respdflist.xlsx"
    base_path = r"C:\Users\Amy19\Desktop\收文"
    basedir = base_path + pd.Timestamp.today().strftime("%Y%m%d")
    # basedir = r"C:\Users\Amy19\Desktop\收文20220419"

    resfn = basedir + "\\" + "respdflist.xlsx"
    resdf = populatedf(basedir, resfn) #################
    highlight_respdflist(resfn) #################

    print("docmark,isattach,recsource,pubdate must be all checked beforehand")
    print("toreceive,toupload are manually added")
    input("After manual check, press to continue...\n")
    resfn = basedir + "\\" + "respdflist_modified.xlsx"
    resdf1 = fill_docmark_forattach(resfn)
    resdf2 = update_fpath(basedir, resfn)
    # fill recno for all files, including those whose "toreceive" set to False
    recfn= r"C:\Users\Amy19\Documents\00000-收发文2022\收文\2022年收文登记簿.xlsx"
    resdf3 = fill_recno(resfn, recfn)

    input("...\n")

