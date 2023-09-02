# -*- coding: utf-8 -*-

#follow order as below:
#rotateToPortrait
#store and add PageNum
#judgeSpecificPageNum then addBlankPage

## TODO enforceColorSpace

import pandas as pd
import os
import tempfile
import re
import sys

import logging
import logging.config
import yaml

with open('config.yaml', 'r') as f:
    config = yaml.safe_load(f.read())
    logging.config.dictConfig(config)
# logger = logging.getLogger(__name__)
logger = logging.getLogger('verboseLogger')
# logger.setLevel(logging.DEBUG)

sys.setrecursionlimit(3500) #3000
#modified C:\Software\Anaconda3\Lib\site-packages\PyPDF2\utils.py
#disable proxy before launching anaconda and using pip

from mergepdf import mergepdf,replaceWithGrayPagespdf
from overlaypdf import overlaypdf
from inspectcolorpage import inspectcolorpage

#pdflist.txt 需要预处理：检查排序
########################## for test use #####################
# pdflistfn="pdflist_send_test.txt"
# flag="发文"

# pdflistfn="pdflist_receive_test.txt"
# flag="收文"
############################################################

# TODO!!! 先修改归档章信息，见 toprint/bulkmergepdf/prepareStamp/generateStamp.py
# pdflistfn="../pdflist_send.txt"
# flag="发文"
toconvertfn="../toconvert.txt"
pdflistfn="../pdflist_receive.txt"
flag="收文"


#res_print_send.xlsx 后续用foldername检索对应归档件号，再写入扉页归档章中
orderfn="res_print.xlsx"
mergedfn="merged_original.pdf"
overlayedfn="overlayed.pdf"

convertedgrayfn="convertedgray.pdf"
toprintfn="toprint.pdf"
pnostampdir="data//pnoStamp//"
coverstampdir="data//coverStamp//"

dffn=re.sub("\.txt","_df.xlsx",pdflistfn,1)
# examine zhengwen_otherpnolist then list unwanted color pages in it 
# moreToConvertPages=[]
year=2021

def loaddf(dffn):
    df=pd.read_excel(dffn,index_col=0)
    df=df.convert_dtypes()
    return df

if not os.path.isdir(pnostampdir):
    os.mkdir(pnostampdir)
if not os.path.isdir(coverstampdir):
    os.mkdir(coverstampdir)

if os.path.isfile(dffn) and os.path.isfile(mergedfn):
    df=loaddf(dffn)
else:
    df=mergepdf(pdflistfn=pdflistfn,orderfn=orderfn,mergedfn=mergedfn)
    df.to_excel(dffn)

if not os.path.isfile(overlayedfn):
    overlaypdf(flag, df,year,pnostampdir,coverstampdir,mergedfn,overlayedfn) # add stamp per page
else:
    logger.info("==================== Exsists overlayed pdf ====================")


# if not os.path.isfile(toconvertfn):
#     toConvertPages = decideToConvertPages(flag, df, overlayedfn, toconvertfn)
# else:
#     logger.info(f"==================== Exists {toconvertfn}, loading toConvertPages ====================")
#     txt = open(toconvertfn, "r", encoding="utf-8")
#     l = txt.readline()
#     toConvertPages = []
#     while l:
#         tmpstr = l.strip()
#         if tmpstr:
#             toConvertPages.append(int(tmpstr))
#         l = txt.readline()
#     txt.close()


# toConvertPages=toConvertPages + moreToConvertPages #######

inspectcolorpage(flag,df,toconvertfn,overlayedfn,convertedgrayfn)

replaceWithGrayPagespdf(toconvertfn,overlayedfn,convertedgrayfn,toprintfn)



