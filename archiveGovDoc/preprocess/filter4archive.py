#!/usr/bin/env python
# coding: utf-8
import os.path

import pandas as pd
from utils.writechswb import writechswb


def filter_sheet(df):
    return df[df['年限'].notnull()]



def filter_wb(fn,resfn):
    df=pd.read_excel(fn,sheet_name=None)
    keys=list(df.keys())
    # resdf=pd.DataFrame()
    # for k in keys[1:]:
    #     tmpdf=df[k]
    #     resdf=resdf.append(tmpdf[tmpdf['年限'].notnull()])
    resdf = pd.concat([filter_sheet(df[k]) for k in keys[1:]], ignore_index=True)

    writechswb(resdf,resfn)
    return resdf


def applyfilter(fn_send, fn_receive, resfn):
    tmpfilename = os.path.basename(fn_send)
    tmpdirname = os.path.dirname(fn_send)
    tmpresfn = os.path.join(tmpdirname, "midres_filtered_"+tmpfilename[5:7]+"_"+tmpfilename[:4]+".xlsx")
    resdf_send=filter_wb(fn_send,tmpresfn)

    tmpfilename = os.path.basename(fn_receive)
    tmpdirname = os.path.dirname(fn_receive)
    tmpresfn = os.path.join(tmpdirname, "midres_filtered_"+tmpfilename[5:7]+"_"+tmpfilename[:4]+".xlsx")
    resdf_receive=filter_wb(fn_receive,tmpresfn)

    resdf_send["备注"]=""
    resdf_receive["备注"]=""
    resdf_receive.rename(columns={"来文字号":"文号","文件标题":"题名"},inplace=True)

    # finalresdf=resdf_send.append(resdf_receive)
    finalresdf = pd.concat([resdf_send, resdf_receive], ignore_index=True)
    finalresdf[["收发文","序列"]]=""
    finalresdf=finalresdf[["时间","文号","收发文","序列","年限","题名","备注"]]

    writechswb(finalresdf, resfn)



if __name__ == "__main__":
    fn_send="raw//2021年发文登记簿.xlsx"
    # fn_send="2021年发文登记簿（原北京）.xlsx"
    fn_receive="raw//2021年收文登记簿.xlsx"

    resfn="raw//empty_排序_2021.xlsx"
    # resfn="raw//empty_排序_2021_beijing.xlsx"

    applyfilter(fn_send, fn_receive, resfn)
