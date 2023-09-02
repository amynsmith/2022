#!/usr/bin/env python
# coding: utf-8

import pandas as pd

import os,sys
# sys.path.append(os.path.abspath(".."))
from utils.writechswb import writechswb
import re

# fn1="页码.xlsx"


# fn2="排序_2021.xlsx" #XX分公司
# fn2="排序_2021_beijing.xlsx" #原XX分公司

# df1=pd.read_excel(fn1)
# df1.set_index(["文号","题名"],inplace=True)
# df2=pd.read_excel(fn2)
# df2.set_index(["文号","题名"],inplace=True)
# df=df2.join(df1,how="outer",rsuffix="s_")
# df.to_excel("midres_pagenum.xlsx")


# string="XX安〔2020〕15号"
# string="第1 号"
# string="〔2021〕第1期"
# int(string.split("〕")[-1].strip("第号期 "))

def extractnum(string):
    num=string.split("〕")[-1].strip("第号期 ")
    return int(num)


def customsort(fn,resfn):
    df = pd.read_excel(fn)
    # df.set_index(["文号","题名"],inplace=True)
    # resdf=df.reset_index()
    df[["已下载", "已打印", "页码"]] = ""
    df["备注"] = ""
    resdf = df[["已下载", "已打印", "页码", "备注", "文号", "题名", "时间", "收发文", "序列", "年限"]].copy()
    resdf["num"] = resdf["文号"].apply(extractnum)
    resdf.sort_values(by=["年限", "收发文", "序列", "num"], inplace=True)
    resdf["文件夹名称"] = resdf["文号"] + " " + resdf["题名"].apply(lambda x: re.sub("\W", "", x))
    # 收文类的时间列设置为" "，而非空值，避免题名列过宽遮挡时间列
    resdf.loc[resdf["收发文"] == "收文", "时间"] = " "
    writechswb(resdf, resfn)




if __name__ == "__main__":
    fn = "sort//tag_排序_2021_beijing.xlsx"
    resfn = "sort//res_sort_beijing.xlsx"

    customsort(fn,resfn)
