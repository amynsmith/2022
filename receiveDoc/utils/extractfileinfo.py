import logging
logger = logging.getLogger('verboseLogger')
# logger.propagate = False  # ineffective

from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
import pdfplumber
import pandas as pd
import os
import tempfile
import re
import numpy as np
import subprocess
import glob
import shutil

exe = "C:\\Software\\gs9.55.0\\bin\\gswin64c.exe"

exclude_ptn = re.compile(r"(^res)|(_resave.pdf$)|(.crdownload$)|(.tmp$)")
title_ptn = "(?<=\d[号期]).+"
docmark_ptn = ".+?\d[号期]"


def search_pubdate_fromtext(text, pageind):
    tmpstr = re.sub("\W", "", text)
    logger.debug(f"\nsearch_pubdate_fromtext...")
    logger.debug(f"tmpstr:\n {tmpstr}")
    # tmpstr="YYXX2022年3月X27日X"
    if "年" in tmpstr:
        # pubdatestr=re.findall("20[1-9][0-9]年{0,1}\d{1,2}月{0,1}\d{1,2}日{0,1}",tmpstr)[pageind]
        # pubdatestr=re.findall("20[1-9][0-9]年.+?日",tmpstr)[pageind]
        tmp = re.findall("20[1-9][0-9]年\w{3,10}日", tmpstr)
        if len(tmp) < 1:
            logger.debug("********** fail to extract date **********")
            pubdate = " "
        else:
            pubdatestr = tmp[pageind]
            pubdate = re.sub("\D", "", pubdatestr)
    else:
        # tmpstr="202211112220000222202762022330"
        tmp = re.findall("20[1-9][0-9]\d{2,4}(?=\D*$)", tmpstr)
        if len(tmp) < 1:
            logger.debug("********** fail to extract date **********")
            pubdate = " "
        else:
            pubdate = tmp[-1]
    logger.debug(f"pubdate:\n {pubdate}")
    return pubdate


def search_docmark_fromtext(text):
    tmpstr = re.sub("[^\w\d〔〕]", "", text)
    logger.debug(f"\nsearch_docmark_fromtext...")
    logger.debug(f"tmpstr:\n {tmpstr}")
    # no need to findall
    searchres = re.search("X.{0,2}Y.{2,10}〔.{4,6}〕.{1,10}号", tmpstr)
    matchstr = ["" if searchres is None else searchres.group()][0]
    logger.debug(f"1st match str: {matchstr}")
    if matchstr != "":
        docmark = ""
        tmpl = re.split("[〔〕]", matchstr)
        tmpl.insert(1, "〔")
        tmpl.insert(3, "〕")
        # ['XX111YY人', '〔', '20ss22', '〕', '1ss46号']
        ptnparts = ["\d", "", "\D", "", "[^\d号]"]
        for ind, tmp in enumerate(tmpl):
            tmpptn = ptnparts[ind]
            docmark = docmark + re.sub(tmpptn, "", tmp)
    else:
        # 第2号 督办通报 第3期
        tmpstr = re.sub("[^\w\d〔〕]", "", text)
        searchres = re.search("第.{1,10}[号期]", tmpstr)
        matchstr = ["" if searchres is None else searchres.group()][0]
        logger.debug(f"2nd match str: {matchstr}")
        docmark = re.sub("^\d[第号期]", "", matchstr)
    logger.debug(f"docmark: {docmark}")
    return docmark


# def populate_pubdatedf(df, df_indslices, page_indlist):
#     # divide into several parts by df_indlist
#     for i, df_indlist in enumerate(df_indslices):
#         page_ind = page_indlist[i]
#         if len(df_indlist) <= 0:
#             continue
#         pdflist = df.iloc[df_indlist].fpath.values
#         tmp_pubdatelist, tmp_pagenumlist = extract_pubdatelist(pdflist, pageind=page_ind)
#         df.loc[df_indlist, "pagenum"] = tmp_pagenumlist
#         df.loc[df_indlist, "pubdate"] = tmp_pubdatelist
#     return df


def identifyissuer(docmark):
    if re.search(r"\AXXYYXX", docmark):
        return "分公司"
    elif re.search(r"\AXXYY", docmark):
        return "X公司"
    elif re.search(r"\AXXY", docmark):
        return "X分局"
    elif re.search(r"\AXXY", docmark):
        return "X局"
    else:
        return ""


def getsheet4receive(no, recsource):
    if (re.search("^XXYY党", no)):
        return "Y公司党委来文"
    elif (re.search("^XXYY纪", no)):
        return "Y公司纪委来文"
    elif (re.search("^XXYY团", no)):
        return "Y公司团委来文"
    elif (re.search("^XXYY工会", no)):
        return "Y公司工会来文"
    elif (re.search("^XXYY", no)):
        return "Y公司行政来文"

    elif (re.search("^XXY党", no)):
        return "X局党委来文"
    elif (re.search("^XXY纪", no)):
        return "X局纪委来文"
    elif (re.search("^XXY团", no)):
        return "X局团委来文"
    ################ XXY群
    elif (re.search("^XXY工会", no)):
        return "X局工会来文"
    ####################
    elif (re.search("^XXYZ", no)):
        return "Y分局来文"
    #######################################
    # XXY工函〔2022〕30号
    elif (re.search("^XXY[^YZ党纪团]函{0,1}", no)):
        return "X局行政来文"
    elif (re.search("^X局集团人", no)):
        return "其他来文"
    else:
        if recsource == "Y公司":
            return "Y公司行政来文"
        elif recsource == "X局":
            return "X局行政来文"
        else:
            return "其他来文"


def get_newname_forattach(fn):
    dirname = os.path.dirname(fn)
    fbase = os.path.basename(fn)
    # fname = os.path.splitext(fbase)[0]
    # ext = os.path.splitext(fbase)[-1]
    if re.search(docmark_ptn, fbase) is not None:
        ori_docmark = re.search(docmark_ptn, fbase).group()
        # print(f"ori_docmark: {ori_docmark}")
        newbase = fbase.replace(ori_docmark, "")
    else:
        newbase = fbase
    if re.search("附件.+", fbase) is not None:
        newbase = "附件" + newbase.split("附件", 1)[-1]
    newbase = newbase.strip()
    newbase = newbase.replace("&nbsp;", "").replace("&nbsp", "")
    # print(f"rename attach file from {fbase} to {newbase}")
    newfn = os.path.join(dirname, newbase)
    return newfn


def get_newname_forbody(fn, docmark):
    dirname = os.path.dirname(fn)
    fbase = os.path.basename(fn)
    # includes docmark==""
    if docmark not in fbase:
        newbase = fbase.strip()
        newbase = docmark + " " + newbase
        newbase = newbase.replace("&nbsp;", "").replace("&nbsp", "")
        # print(f"rename body file from {fbase} to {newbase}")
        newfn = os.path.join(dirname, newbase)
    else:
        newbase = fbase.strip()
        newbase = newbase.replace("&nbsp;", "").replace("&nbsp", "")
        newfn = os.path.join(dirname, newbase)
    newfn = newfn.replace("&nbsp;", "").replace("&nbsp", "")
    return newfn


def rename_attachfile(fn):
    newfn = get_newname_forattach(fn)
    if newfn != fn:
        shutil.move(fn, newfn)
    if fn.endswith("_resave.pdf"):
        orifn = fn.rsplit("_resave.pdf", 1)[0] + ".pdf"
        newfn = get_newname_forattach(orifn)
        if newfn != orifn:
            shutil.move(orifn, newfn)
    tmp = os.path.basename(newfn)
    title = os.path.splitext(tmp)[0]
    return title, newfn


def rename_bodyfile(fn, docmark):
    newfn = get_newname_forbody(fn, docmark)
    if newfn != fn:
        shutil.move(fn, newfn)
    if fn.endswith("_resave.pdf"):
        orifn = fn.rsplit("_resave.pdf", 1)[0] + ".pdf"
        newfn = get_newname_forbody(orifn, docmark)
        if newfn != orifn:
            shutil.move(orifn, newfn)
    tmp = os.path.basename(newfn)
    title = os.path.splitext(tmp)[0]
    if docmark in title:
        title = title.replace(docmark, "").strip()
    return title, newfn


def extract_fileinfo(fn):
    # ["docmark", "isattach", "title", "recsource", "fpath"]
    # pagenum,pubdate
    dirname = os.path.dirname(fn)
    foldername = os.path.split(dirname)[-1]
    fbase = os.path.basename(fn)
    fname = os.path.splitext(fbase)[0]
    ext = os.path.splitext(fbase)[-1]
    logger.debug(f"current folder: {foldername}")

    filename_docmark = ["" if re.search(docmark_ptn, fbase) is None else re.search(docmark_ptn, fbase).group()][0]

    if ext != ".pdf":
        docmark = ""
        isattach = True
        recsource = ""
        pagenum = 0
        pubdate = ""
        title, fpath = rename_attachfile(fn)
        logger.info(f"{docmark}\t {title}\t isattach:{isattach}\t {recsource}\t {pubdate}")
        return docmark, isattach, title, recsource, fpath, pagenum, pubdate

    f = open(fn, 'rb')
    reader = PdfFileReader(f)
    if reader.isEncrypted:
        f.close()
        logger.debug("**********is encrypted pdf**********")
        pdfinputfn = fn
        pdfoutputfn = fn.split(".pdf")[0] + "_resave.pdf"
        if not os.path.isfile(pdfoutputfn):
            logger.debug("**********resave needed**********")
            subprocess.run([exe, "-o", pdfoutputfn, "-sDEVICE=pdfwrite", pdfinputfn])
        # replace with the new resaved pdf
        fn = pdfoutputfn
    if f.closed is False:
        f.close()
    # now decrypted

    with pdfplumber.open(fn) as pdf:
        page = pdf.pages[0]
        text = page.extract_text()
    # using PdfFileReader not effective
    docmark = search_docmark_fromtext(text)

    with open(fn, 'rb') as f:
        reader = PdfFileReader(f)
        pagenum = reader.getNumPages()

    rootpath = dirname + "\\*.*"
    # folderfiles = list(glob.iglob(rootpath))
    tmpl = list(glob.iglob(rootpath))
    folderfiles = [i for i in tmpl if exclude_ptn.search(os.path.basename(i)) is None]
    if docmark == "":
        recsource = ""
        pubdate = ""
        # last rename
        if len(folderfiles) > 1:
            isattach = True
            title, fpath = rename_attachfile(fn)
        else:
            # the only one file is body file
            isattach = False
            title, fpath = rename_bodyfile(fn, docmark)
        logger.info(f"{docmark}\t {title}\t isattach:{isattach}\t {recsource}\t {pubdate}")
        return docmark, isattach, title, recsource, fpath, pagenum, pubdate
    elif filename_docmark not in ["", docmark] and len(folderfiles) > 1:
        # extracted docmark in pdf content
        # perhaps is attach file which references another docmark
        recsource = ""
        pubdate = ""
        isattach = True
        docmark = ""
        title, fpath = rename_attachfile(fn)
        logger.info(f"{docmark}\t {title}\t isattach:{isattach}\t {recsource}\t {pubdate}")
        return docmark, isattach, title, recsource, fpath, pagenum, pubdate
    else:
        # determine to be body file
        isattach = False
        recsource = identifyissuer(docmark)
        # determine pubdate search loc
        pageind = [0 if docmark.endswith("期") else -1][0]

        with open(fn, 'rb') as f:
            reader = PdfFileReader(f)
            text = reader.getPage(pageind).extractText()

        tmpstr = re.sub("\W", "", text)
        if tmpstr != "" and len(re.findall("20[1-9][0-9]\d{2,4}(?=\D*$)", tmpstr)) > 0 and pageind == -1:
            # prefer searching pubdate from the last page
            pubdate = search_pubdate_fromtext(text, pageind)
        else:
            # pageind == 0
            with pdfplumber.open(fn) as pdf:
                page = pdf.pages[pageind]
                text = page.extract_text()
                pubdate = search_pubdate_fromtext(text, pageind)
        # last rename
        title, fpath = rename_bodyfile(fn, docmark)
        logger.info(f"{docmark}\t {title}\t isattach:{isattach}\t {recsource}\t {pubdate}")
        return docmark, isattach, title, recsource, fpath, pagenum, pubdate


if __name__ == "__main__":
    # fn1 = "..\\项目章清理\\2018年以来刻制项目章清单_局.xlsx"
    # fn2 = "..\\项目章清理\\2018年以来刻制项目章清单_公司.xlsx"
    #
    # resfn1 = "..\\项目章清理\\2018年以来刻制项目章清单_局_pubdate.xlsx"
    # resfn2 = "..\\项目章清理\\2018年以来刻制项目章清单_公司_pubdate.xlsx"
    #
    # df1 = pd.read_excel(fn1)
    # df2 = pd.read_excel(fn2)
    #
    # pdflist1 = df1.fpath.values
    # pdflist2 = df2.fpath.values
    #
    # pubdatelist1, tmppagenumlist = extract_pubdatelist(pdflist1)
    # pubdatelist2, tmppagenumlist = extract_pubdatelist(pdflist2)
    #
    # resdf1 = pd.DataFrame(data=np.column_stack([pdflist1, pubdatelist1]), columns=["fn", "pubdate"])
    # resdf2 = pd.DataFrame(data=np.column_stack([pdflist2, pubdatelist2]), columns=["fn", "pubdate"])
    # resdf1.pubdate = resdf1.pubdate.apply(lambda s: pd.to_datetime(s, format="%Y%m%d"))
    # resdf2.pubdate = resdf2.pubdate.apply(lambda s: pd.to_datetime(s, format="%Y%m%d"))
    #
    # resdf1.to_excel(resfn1)
    # resdf2.to_excel(resfn2)

    basedir = r"C:\Users\Amy19\Desktop\新建文件夹"
    rootpath = basedir + r"\*\*.pdf"
    for i in glob.iglob(rootpath):
        dirname = os.path.dirname(i)
        foldername = os.path.split(dirname)[-1]
        print(foldername)
        print(os.path.basename(i))
