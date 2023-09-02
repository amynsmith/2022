import pandas as pd
import os
import sys

sys.path.append(os.path.abspath(".."))
sys.path.append(r"C:\Users\Amy19\PycharmProjects\receiveDoc\utils")
import re

from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
import pdfplumber

import subprocess

exe = "C:\\Software\\gs9.55.0\\bin\\gswin64c.exe"

from extractfileinfo import search_pubdate_fromtext


# 档案系统上传文件类型
def getftype(fname):
    if fname == "正文.pdf":
        return "正文"
    elif fname in ["文单.pdf", "处理记录.pdf"]:
        return "处理单"
    else:
        return "附件"


def getfpath(basepath, row):
    if os.path.basename(basepath) == "发文" or os.path.basename(basepath) == "发文（原北京）":
        if row.pdf == "正文.pdf":
            fpath = os.path.join(basepath, row.title, row.title + ".pdf")  # 发文类上传盖有电子章的原版，而非“正文.pdf”
        else:
            fpath = os.path.join(basepath, row.title, row.pdf)
    else:
        if row.title.startswith("XX京"):
            # basepath = basepath.rsplit("收文\\\\",1)[0]+"发文（原北京）\\\\"
            basepath = os.path.join(os.path.dirname(basepath), "发文（原北京）")
        elif row.title.startswith("XX北方"):
            # basepath = basepath.rsplit("收文\\\\",1)[0]+"发文\\\\"
            basepath = os.path.join(os.path.dirname(basepath), "发文")
        else:
            basepath = basepath
        fpath = os.path.join(basepath, row.title, row.pdf)
    return fpath


def fillftype(basepath, df):
    if "发文" in os.path.basename(basepath):
        df["ftype"] = df.pdf.apply(getftype)
        return df
    else:
        # 收文
        for i in range(len(df)):
            if i > 0 and df.loc[i - 1, "pdf"] == "文单.pdf":
                if df.loc[i - 1, "stamp_order"] > 0:
                    df.loc[i, "ftype"] = "正文"
                else:
                    # 批复对应的请示件（正文、文单、处理记录、定稿）
                    df.loc[i, "ftype"] = "附件"
            elif df.loc[i, "pdf"] in ["文单.pdf", "处理记录.pdf"]:
                if re.search(r"归档资料—打印文件夹\\收文.*", df.loc[i, "fpath"]) is None:
                    # 批复对应的请示件（正文、文单、处理记录、定稿）
                    df.loc[i, "ftype"] = "附件"
                else:
                    df.loc[i, "ftype"] = "处理单"
            else:
                df.loc[i, "ftype"] = "附件"
        return df


def generatedf(fn, resfn, archiveno_base, basepath):
    df = pd.read_excel(fn)
    # 档案系统上传文件路径
    # df["fpath"] = df.apply(lambda row: os.path.join(basepath, row.title, row.pdf), axis=1)
    df["fpath"] = df.apply(lambda row: getfpath(basepath, row), axis=1)

    titleset = set(df.title)
    # 档号
    # 10年 0303.0600-A-2021-D-0001
    # 30年 0303.0600-A-2021-C-0001
    df["archive_no"] = ""
    for s in titleset:
        tmpdf = df[df.title == s]
        tmpind = tmpdf.index
        # stamp_order = tmpdf[tmpdf.pdf == "正文.pdf"].stamp_order.values[0]
        stamp_order = max(tmpdf.stamp_order.values)
        archive_no = archiveno_base + str(stamp_order).zfill(4)
        df.loc[tmpind, "archive_no"] = archive_no
    # 档案系统上传文件类型
    # df["ftype"] = df.pdf.apply(getftype)
    df = fillftype(basepath, df)
    # 对收文中被合并的请示件，其档号应用对应收文的档号，而非形如 0303.0600-A-2021-C-0000
    for s in titleset:
        tmpdf = df[df.title == s]
        tmpind = tmpdf.index
        if max(tmpdf.stamp_order.values) == 0:
            print(f"archive_no of {s} should be corrected...")
            prevind=min(tmpind)-1
            while prevind >0:
                prev_ftype = df.loc[prevind, "ftype"]
                if prev_ftype == "处理单":
                    stamp_order=df.loc[prevind, "stamp_order"]
                    print(f"stamp_order:{stamp_order}")
                    archive_no = archiveno_base + str(stamp_order).zfill(4)
                    df.loc[tmpind, "archive_no"] = archive_no
                    break
                else:
                    prevind = prevind -1
    resdf = df[["title", "pdf", "stamp_order", "archive_no", "ftype", "fpath"]]
    resdf.to_excel(resfn)


def addpubdate(fn):
    df = pd.read_excel(fn)
    bodydf = df[df.ftype == "正文"]
    bodyind = bodydf.index
    pageind = -1
    df["pubdate"] = ""
    for ind in bodyind:
        print(f"cur ind: {ind}")
        fpath = df.loc[ind, "fpath"]
        # ensure pdf to be decrypted
        pdfoutputfn = fpath.split(".pdf")[0] + "_resave.pdf"
        if os.path.isfile(pdfoutputfn):
            fpath = pdfoutputfn
        else:
            pdfinputfn = fpath
            subprocess.run([exe, "-o", pdfoutputfn, "-sDEVICE=pdfwrite", pdfinputfn])
            fpath = pdfoutputfn
        # now decrypted
        with open(fpath, 'rb') as f:
            reader = PdfFileReader(f)
            text = reader.getPage(pageind).extractText()
        tmpstr = re.sub("\W", "", text)
        if tmpstr != "" and len(re.findall("20[1-9][0-9]\d{2,4}(?=\D*$)", tmpstr)) > 0 and pageind == -1:
            # prefer searching pubdate from the last page
            pubdate = search_pubdate_fromtext(text, pageind)
        else:
            with pdfplumber.open(fpath) as pdf:
                page = pdf.pages[pageind]
                text = page.extract_text()
                pubdate = search_pubdate_fromtext(text, pageind)
        df.loc[ind, "pubdate"] = pubdate
        print(f"added pubdate: {pubdate}")
    df.loc[:, "pubdate"] = pd.to_datetime(df.pubdate, format="%Y%m%d", errors="coerce")
    df.to_excel(fn, index=False)


if __name__ == "__main__":
    # 10年 0303.0600-A-2021-D-0001
    # 30年 0303.0600-A-2021-C-0001
    # C:\Users\Amy19\PycharmProjects\archiveGovDoc\toprint\bulkmergepdf\res 2021发文（北方）
    #
    # fn = r"C:\Users\Amy19\Desktop\文书归档-合并pdf打印用 bak\bulkmergepdf\res 2021发文（北方）\pdflist_send_df.xlsx"
    # resfn = r"C:\Users\Amy19\PycharmProjects\archiveGovDoc\upload\C0026-0067_30年发文（北方）_toupload.xlsx"
    # archiveno_base = "0303.0600-A-2021-C-"
    # basepath = "C:\\Users\\Amy19\\Documents\\0000-文书归档\\2021\\归档资料—打印文件夹\\发文"
    #
    # fn = r"C:\Users\Amy19\Desktop\文书归档-合并pdf打印用 bak\bulkmergepdf\res 2021收文（北方）\pdflist_receive_df.xlsx"
    # resfn = r"C:\Users\Amy19\PycharmProjects\archiveGovDoc\upload\C0068-0157_30年收文_toupload.xlsx"
    # archiveno_base = "0303.0600-A-2021-C-"
    # basepath = "C:\\Users\\Amy19\\Documents\\0000-文书归档\\2021\\归档资料—打印文件夹\\收文"
    #
    # fn = r"C:\Users\Amy19\PycharmProjects\archiveGovDoc\toprint\bulkmergepdf\res 2021发文（北京）30年\pdflist_send_df.xlsx"
    # resfn = r"C:\Users\Amy19\PycharmProjects\archiveGovDoc\upload\C0001-0025_30年发文（北京）_toupload.xlsx"
    # archiveno_base = "0303.0600-A-2021-C-"
    # basepath = "C:\\Users\\Amy19\\Documents\\0000-文书归档\\2021\\归档资料—打印文件夹\\发文（原北京）"
    #
    # fn = r"C:\Users\Amy19\PycharmProjects\archiveGovDoc\toprint\bulkmergepdf\res 2021发文（北京）10年\pdflist_send_df.xlsx"
    # resfn = r"C:\Users\Amy19\PycharmProjects\archiveGovDoc\upload\D0001-0101_10年发文（北京）_toupload.xlsx"
    # archiveno_base = "0303.0600-A-2021-D-"
    # basepath = "C:\\Users\\Amy19\\Documents\\0000-文书归档\\2021\\归档资料—打印文件夹\\发文（原北京）"
    #
    # fn = r"C:\Users\Amy19\PycharmProjects\archiveGovDoc\toprint\bulkmergepdf\res 2021发文（北方）10年\pdflist_send_df.xlsx"
    # resfn = r"C:\Users\Amy19\PycharmProjects\archiveGovDoc\upload\D0102-0162_10年发文（北方）_toupload.xlsx"
    # archiveno_base = "0303.0600-A-2021-D-"
    # basepath = "C:\\Users\\Amy19\\Documents\\0000-文书归档\\2021\\归档资料—打印文件夹\\发文"
    #
    # fn = r"C:\Users\Amy19\PycharmProjects\archiveGovDoc\toprint\bulkmergepdf\res 2021收文（北方）10年\pdflist_receive_df.xlsx"
    # resfn = r"C:\Users\Amy19\PycharmProjects\archiveGovDoc\upload\D0163-0170_10年收文（北方）_toupload.xlsx"
    # archiveno_base = "0303.0600-A-2021-D-"
    # basepath = "C:\\Users\\Amy19\\Documents\\0000-文书归档\\2021\\归档资料—打印文件夹\\收文"
    #
    # generatedf(fn, resfn, archiveno_base, basepath)
    # 仅收文需要添加”成文日期“列
    # addpubdate(resfn)

    # allfnlist = [r"C:\Users\Amy19\PycharmProjects\archiveGovDoc\upload\C0001-0025_30年发文（北京）_toupload.xlsx",
    #           r"C:\Users\Amy19\PycharmProjects\archiveGovDoc\upload\C0026-0067_30年发文（北方）_toupload.xlsx",
    #           r"C:\Users\Amy19\PycharmProjects\archiveGovDoc\upload\D0163-0170_10年收文（北方）_toupload.xlsx",
    #           r"C:\Users\Amy19\PycharmProjects\archiveGovDoc\upload\C0068-0157_30年收文_toupload.xlsx",
    #           r"C:\Users\Amy19\PycharmProjects\archiveGovDoc\upload\D0171-0175_10年其他_toupload.xlsx",
    #           r"C:\Users\Amy19\PycharmProjects\archiveGovDoc\upload\D0102-0162_10年发文（北方）_toupload.xlsx",
    #           r"C:\Users\Amy19\PycharmProjects\archiveGovDoc\upload\D0001-0101_10年发文（北京）_toupload.xlsx"]
    #
    # resfn = r"C:\Users\Amy19\PycharmProjects\archiveGovDoc\upload\toupload.xlsx"
    # resdf = pd.concat([pd.read_excel(fn) for fn in allfnlist],ignore_index=True)
    # resdf["toupload"]=True
    # resdf.to_excel(resfn)

    fn = r"C:\Users\Amy19\PycharmProjects\archiveGovDoc\upload\toupload.xlsx"
    df=pd.read_excel(fn)
    fpathlist = df.fpath.values
    for fpath in fpathlist:
        if not os.path.isfile(fpath):
            print(fpath)

    print(len(set(df.archive_no))) # 332