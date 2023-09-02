import pandas as pd
import os,sys
sys.path.append(os.path.abspath(".."))
from utils.writechswb import writechswb
import re


def addtag4send(no,title):  
    if (re.search("办公例会.*会议纪要\Z",title)):
        return "纪要（例会）"
    elif (re.search("领导班子会.*会议纪要\Z",title)):
        return "纪要（班子会）"
    elif (re.search("党委会.*会议纪要\Z",title)):
        return "纪要（党）"
    elif (re.search("第[0-9]+期\Z",no) and re.search("督办事项.*通报\Z",title)):
        return "督办通报"
    elif (re.search("\AXX(北方|京)纪",no)):
        return "纪委"
    elif (re.search("\AXX(北方|京)团",no)):
        return "团委"
    elif (re.search("\AXX(北方|京)工会",no)):
        return "工会"
    elif (re.search("\AXX(北方|京)党",no)):
        if ("函" in no):
            return "党委函"
        else:
            return "党委"
    elif ("函" in no):
        return "行政函"
    else:
        return "行政"

def addtag4receive(no,title):
    if (re.search("\AXX党",no)):
        if ("函" in no):
            return "公司党委函"
        else:
            return "公司党委"
    elif (re.search("\AXX纪",no)):
        return "公司纪委"       
    elif (re.search("\AXX团",no)):
        return "公司团委"
    elif (re.search("\AXX工会",no)):
        return "公司工会"
    elif (re.search("\AXX",no)):
        if ("函" in no):
            return "公司行政函"
        else:
            return "公司行政"
    elif (re.search("\AXX党",no)):
        if ("函" in no):
            return "局党委函"
        else:
            return "局党委"
    elif (re.search("\AXX[^纪团(工会)]",no)):
        if ("函" in no):
            return "局行政函"
        else:
            return "局行政"
    elif (re.search("\AX局集团人",no)):
        return "集团行政"
    else:
        return ""
    

def funcwrap(row):
    no=row["文号"]
    title=row["题名"]
    if (row["收发文"]=="发文"):
        return addtag4send(no,title)
    else:
        return addtag4receive(no,title)


def addtag(fn,resfn):
    df = pd.read_excel(fn)
    df.loc[df["时间"].isnull(), "收发文"] = "收文"
    df.loc[~df["时间"].isnull(), "收发文"] = "发文"
    df["序列"]=df.apply(funcwrap,axis=1)
    writechswb(df,resfn)




if __name__ == "__main__":
    # fn="raw//empty_排序_2021.xlsx"
    # resfn = "sort//tag_排序_2021.xlsx"
    fn = "raw//empty_排序_2021_beijing.xlsx"
    resfn = "sort//tag_排序_2021_beijing.xlsx"

    addtag(fn, resfn)
