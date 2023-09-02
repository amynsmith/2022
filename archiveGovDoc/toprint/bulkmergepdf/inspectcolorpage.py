import re
import re
import subprocess
import math
import numpy as np
import os

from mergepdf import simple_mergepdf
from common import writePagenumList, loadPagenumList

import logging
import logging.config
import yaml
# logging.basicConfig(level=logging.ERROR)
with open('config.yaml', 'r') as f:
    config = yaml.safe_load(f.read())
    logging.config.dictConfig(config)
# logger = logging.getLogger(__name__)
logger = logging.getLogger('verboseLogger')
# logger.setLevel(logging.DEBUG)

RE_FLOAT = re.compile("[01].[0-9]+")
CMYK_NCOLORS = 4
exe="C:\\Software\\gs9.55.0\\bin\\gswin64c.exe"


def checkCMYKperPage(pdfinputfn):
    gs_inkcov = subprocess.Popen([exe, "-o", "-", "-sDEVICE=inkcov", pdfinputfn], stdout=subprocess.PIPE)
    for raw_line in iter(gs_inkcov.stdout.readline, b''):
        line = raw_line.decode('utf8').rstrip()
        fields = line.split()
        if (len(fields) >= CMYK_NCOLORS and all(RE_FLOAT.match(fields[i]) for i in range(CMYK_NCOLORS))):
          cmyk = tuple(float(value) for value in fields[0:CMYK_NCOLORS])
          yield cmyk

def isColorPage(c, m, y, k):
    return c > 0 or m > 0 or y > 0

# def cmyk_to_rgb(c ,m, y, k, cmyk_scale=1, rgb_scale=255):
#     r=round(rgb_scale*(1.0-c/float(cmyk_scale))*(1.0-k/float(cmyk_scale)))
#     g=round(rgb_scale*(1.0-m/float(cmyk_scale))*(1.0-k/float(cmyk_scale)))
#     b=round(rgb_scale*(1.0-y/float(cmyk_scale))*(1.0-k/float(cmyk_scale)))
#     return (r,g,b)

def findColorPages(pdfinputfn):
    for n, pagecmyk in enumerate(checkCMYKperPage(pdfinputfn), start=1):
        if isColorPage(*pagecmyk):
            # logger.debug(f'{(n,pagecmyk)}')
            yield (n, pagecmyk)




def decideToConvertPages(flag,df,inputpdffn,toconvertfn):
    totalColorPages=[i for (i,cmyk) in list(findColorPages(inputpdffn))]
    logger.info("==================== Color pages found: ====================")
    # logger.info(f'{list(findColorPages(inputpdffn))}')
    logger.info(f'{totalColorPages}')
    if flag=="发文":
        #发文类正文，只需彩打首页
        neededColorPages=[int(i) for i in list(df.loc[(df.pdf=="文单.pdf") | (df.pdf=="正文.pdf") , "start_pagenum"])]
        toConvertPages=list(set(totalColorPages)-set(neededColorPages))
    else:
        # flag=="收文"
        #收文类正文，需要彩打正文首页及有印章页，其中被合并的请示件不彩打
        wendan_indlist=df.loc[df.stamp_pagecount>0].index.values
        wendan_pnolist=[int(i) for i in list(df.loc[df.stamp_pagecount>0 , "start_pagenum"].values)]
        #收文类正文紧跟其文单
        # 对于收文类正文，首页均彩打，若后续页面中有图像（即电子章），则彩打
        # 在main函数中读取并追加正文后续页中不含电子章的彩页，再转换
        zhengwen_startpnolist=[]
        zhengwen_otherpnolist=[]
        for i in wendan_indlist:
            # 2022.3.18 updated: 收文不再需要打印处理记录
            # tmp=range(int(df.loc[i+1].start_pagenum), int(df.loc[i+2].start_pagenum))
            tmp=range(int(df.loc[i+1].start_pagenum), int(df.loc[i+1].start_pagenum + df.loc[i+1].pdf_pagecount))
            zhengwen_startpnolist.append(tmp[0])
            zhengwen_otherpnolist.extend(tmp[1:])
        logger.info('==================== pno listed below, further examine needed to extend toConvertPages: ====================')
        logger.info(f'{zhengwen_otherpnolist}')
        tmp_colorpage_list=wendan_pnolist+zhengwen_startpnolist+zhengwen_otherpnolist
        preserveColorPages=list(set(tmp_colorpage_list).intersection(set(totalColorPages)))      
        toConvertPages=list(set(totalColorPages)-set(preserveColorPages))
    # toConvertPages.sort()
    logger.info('==================== Finished deciding toConvertPages: ====================')
    pagenumlist = writePagenumList(toConvertPages, toconvertfn)
    return pagenumlist

#gswin64c -o "convertedgray.pdf" -sDEVICE=pdfwrite -dPDFX=/True -dAutoRotatePages=/None -sPageList="8,18,29,30,31,32,33,34,35,36,38" -sColorConversionStrategy=Gray -sProcessColorModel=DeviceGray "overlayed.pdf" 

def convertColorPages(pdfinputfn,pagenumstr,pdfoutputfn):
    subprocess.run([exe, "-o",pdfoutputfn,"-sDEVICE=pdfwrite", "-dPDFX=/True", "-dAutoRotatePages=/None", "-sPageList="+pagenumstr,
        "-sColorConversionStrategy=Gray","-sProcessColorModel=DeviceGray",pdfinputfn])


def inspectcolorpage(flag,df,toconvertfn,pdfinputfn,pdfoutputfn):
    # pagenumlist=",".join([str(i) for i in toConvertPages])
    # convertColorPages(pdfinputfn,pagenumlist,pdfoutputfn)
    if not os.path.isfile(toconvertfn):
        pagenumlist = decideToConvertPages(flag, df, pdfinputfn, toconvertfn)
    else:
        logger.info(f"==================== Exists {toconvertfn}, loading toConvertPages ====================")
        pagenumlist = loadPagenumList(toconvertfn)
    logger.info(f"==================== Total Chunks Num: {len(pagenumlist)} ====================")
    list_pdfoutputfn=[]
    for ind, tmp_numstr in enumerate(pagenumlist):
        logger.info(f'----> current chunk ind: {ind}')
        tmp_pdfoutputfn=os.path.splitext(pdfoutputfn)[0]+"_chunk"+str(ind)+".pdf"
        if not os.path.isfile(tmp_pdfoutputfn):
            convertColorPages(pdfinputfn, tmp_numstr, tmp_pdfoutputfn)
        else:
            logger.info(f"exists {tmp_pdfoutputfn}")
        list_pdfoutputfn = list_pdfoutputfn + [tmp_pdfoutputfn]
    simple_mergepdf(list_pdfoutputfn, pdfoutputfn)
    logger.info("==================== finished converting color pages ====================")
    return
