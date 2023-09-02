import pandas as pd
import numpy as np
import os
import re
import glob
import os, sys
from datetime import datetime
import shutil
import xlsxwriter
from openpyxl.utils import get_column_letter

def moveto_targetdir(basedir,fn,target_basedir):
    df=pd.read_excel(fn)
    df = df[~df.isattach]
    rootpath = basedir + "\\*\\"
    folderlist = list(glob.iglob(rootpath))
    df["srcdir"]=df.fpath.apply(lambda s: os.path.dirname(s))
    df["foldername"]= df.srcdir.apply(lambda s: os.path.split(s)[-1])
    df["targetdir"]=df[["sheet","foldername"]].apply(lambda row: os.path.join(target_basedir, row.sheet, row.foldername), axis=1)
    resdf = df[["srcdir","targetdir"]]
    for i in range(len(resdf)):
        tmpsrc = resdf.iloc[i].srcdir
        tmptarget = resdf.iloc[i].targetdir
        # print(f"move {tmpsrc} to {tmptarget}")
        shutil.move(tmpsrc, tmptarget)
    return resdf


def str_len(s):
    try:
        row_l=len(s)
        utf8_l=len(s.encode('utf-8'))
        return (utf8_l-row_l)/2+row_l
    except:
        return None



def write_multiplesheets(sheets, dflist, fn):
    Excelwriter = pd.ExcelWriter(fn, engine="xlsxwriter")
    for i, sheetdf in enumerate(dflist):
        sheetdf.to_excel(Excelwriter, sheet_name=sheets[i], index=False)
        worksheet = Excelwriter.sheets[sheets[i]]  # pull worksheet object
        for idx, col in enumerate(sheetdf):  # loop through all columns
            series = sheetdf[col]
            max_len = max((
                series.astype(str).map(str_len).max(),  # len of largest item
                str_len(str(series.name))  # len of column name/header
            )) + 2  # adding a little extra space
            len_limit = 50
            use_len = min(max_len, len_limit)
            worksheet.set_column(idx, idx, use_len)  # set column width
    Excelwriter.save()
    print("multiple sheets save complete")


def write_nibansheet(processtype, df, fn):
    if not os.path.isfile(fn):
        print("new niban file must be created")
        pd.DataFrame().to_excel(fn, sheet_name=processtype)
    else:
        print("niban file exists, now updating it")
    writer = pd.ExcelWriter(fn, engine="openpyxl", mode="a", if_sheet_exists="replace")
    df.to_excel(writer, sheet_name=processtype, index=False)
    # expand cols width
    ws = writer.sheets[processtype]  # pull worksheet object
    targetcolletter = []
    targetwidth=[]
    for i in range(len(df.columns)):
        if df.columns[i] in ["title","href","org","starttime","unknown"]:
            # skip these cols, no need to expand width
            continue
        targetcolletter.append(get_column_letter(i+1))
        series = df.iloc[:,i]
        max_len = max((
            series.astype(str).map(str_len).max(),  # len of largest item
            str_len(str(series.name))  # len of column name/header
        )) + 2  # adding a little extra space
        len_limit = 80
        use_len = min(max_len, len_limit)
        targetwidth.append(use_len)
    for i,colletter in enumerate(targetcolletter):
        ws.column_dimensions[colletter].width = targetwidth[i]  # set column width
    writer.save()




def update_record(fn,recfn):
    cols=["收文\n顺序号","收文日期","来文字号","文件标题"]
    df=pd.read_excel(fn)
    df = df[~df.isattach]
    df = df[["sheet","recno","docmark","title"]]
    df["recorder"] = df.recno.apply(lambda x: int(x.split("-")[-1]))
    modified_sheets = set(df.sheet.values)
    recorddf = pd.read_excel(recfn, sheet_name=None)
    keys = recorddf.keys()
    all_sheetsdf = []
    for k in keys:
        sheetdf = recorddf[k]
        print(f"current sheet {k}")
        if k not in modified_sheets:
            print("direct copy")
            all_sheetsdf = all_sheetsdf +[sheetdf]
        else:
            print("modified, need update")
            if str(sheetdf.iloc[0]["文件标题"]) == 'nan':
                # no previous record
                start_ind = 0
            else:
                # exists previous record
                sheetdf.dropna(axis=0, subset="文件标题", inplace=True)
                start_ind = len(sheetdf)
            # get append values
            tmpdf=df[df.sheet == k]
            #  7-100 ... 7-106 7-99 wrong order
            tmpdf = tmpdf.sort_values(by="recorder")  # don't modify original df
            for i in range(len(tmpdf)):
                print(f"append row at ind: {start_ind}")
                tmpdata = [tmpdf.iloc[i].recno, pd.Timestamp.today().strftime("%Y.%m.%d"),
                            tmpdf.iloc[i].docmark, tmpdf.iloc[i].title]
                tmpdict = dict(zip(cols,tmpdata))
                sheetdf.loc[start_ind]=tmpdict
                start_ind = start_ind +1
            print(f"{len(tmpdf)} rows appended")
            all_sheetsdf = all_sheetsdf +[sheetdf]
    write_multiplesheets(list(keys), all_sheetsdf, recfn)
    return all_sheetsdf





if __name__ == "__main__":
    base_path = r"C:\Users\Amy19\Desktop\收文"
    basedir = base_path + pd.Timestamp.today().strftime("%Y%m%d")
    # basedir = r"C:\Users\Amy19\Desktop\收文20220509"
    target_basedir = r"C:\Users\Amy19\Documents\00000-收发文2022\收文"
    fn = basedir + "\\" + "respdflist_modified.xlsx"
    recfn = r"C:\Users\Amy19\Documents\00000-收发文2022\收文\2022年收文登记簿.xlsx"

    resdf = moveto_targetdir(basedir,fn,target_basedir)
    update_record(fn, recfn)

