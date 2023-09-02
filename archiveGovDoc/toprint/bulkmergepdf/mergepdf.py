import os
import re
import subprocess
import tempfile

import pandas as pd
from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
from reportlab.lib import pagesizes

from common import loadPagenums

exe="C:\\Software\\gs9.55.0\\bin\\gswin64c.exe"

import logging.config
import yaml


with open('config.yaml', 'r') as f:
    config = yaml.safe_load(f.read())
    logging.config.dictConfig(config)
# logger = logging.getLogger(__name__)
logger = logging.getLogger('verboseLogger')
# logger.setLevel(logging.DEBUG)

def isFixedOddPage(index,fn,prevfn):
    foldername=fn.split("\\\\")[-2]
    pdfname=fn.split("\\\\")[-1]
    if (fn.endswith("正文.pdf") or fn.endswith("文单.pdf")):
        return True
    #收文类打印中的正文，含收文正文、有批复故被合并的请示正文
    elif re.sub("\W","",foldername) in re.sub("\W","",pdfname):
        return True
    # 收文的处理记录未紧跟文单，故需要单独在奇数页
    # 2022.3.18 updated: 收文不再需要打印处理记录
    elif (fn.endswith("处理记录.pdf") and not prevfn.endswith("文单.pdf")):
        return True
    else:
        return False


def simple_mergepdf(fnlist, mergedfn):
    simple_merger = PdfFileMerger(strict=False)
    for fn in fnlist:
        reader = PdfFileReader(open(fn, 'rb'))
        simple_merger.append(reader)
    simple_merger.write(mergedfn)


def mergepdf(pdflistfn="pdflist.txt",orderfn="res_print.xlsx",mergedfn="merged_original.pdf"):
    txt = open(pdflistfn, "r",encoding="utf8")
    l=txt.readline()
    fnlist=[]
    foldername=[]
    pdfname=[]
    while l:
        #ignore blank lines represented by "\n"
        if l.strip():
            fnlist.append(l.strip())
            foldername.append(l.strip().split("\\\\")[-2])
            pdfname.append(l.strip().split("\\\\")[-1])
        l=txt.readline()
    txt.close()

    tmp_subfoldername=fnlist[0].rsplit("\\\\",3)[-3] #eg "收文xxx"
    flag=tmp_subfoldername[:2]

    df=pd.DataFrame(data={"title":foldername,"pdf":pdfname})
    df["start_pagenum"]=0
    df["pdf_pagecount"]=0
    df["pdf_append_blankpage"]=0 #assign 0 or 1
    df["stamp_pagecount"]=0
    df["subtract_pagecount"]=0
    df["stamp_order"]=0

    # PdfFileMerger(strict=True) #############
    merger = PdfFileMerger(strict=False)
    with open(mergedfn, "wb") as outputhandle:
        tempPDF = tempfile.TemporaryFile()
        writer = PdfFileWriter()
        writer.addBlankPage(pagesizes.A4[0],pagesizes.A4[1])
        writer.write(tempPDF)

        logger.info("==================== Start merging pdfs ====================")
        for index,fn in enumerate(fnlist):
            # F-string expression part cannot include a backslash
            # logger.debug(f'index {index}')
            logger.debug("index %d : %s", index, fn.split("\\\\")[-1])

    ####################### already handled in preprocess #################################
            # incase: tmp_pdfname='XX审〔2021〕60号  B－L14、B－L15＃、C区C－L1～7＃、C－G1＃、C－G2＃、C－1'
            # ghostscript error: /undefinedfilename
            tmp_pdfname=fn.split("\\\\")[-1].rsplit(".pdf",1)[0]
            # contains " "
            if len(tmp_pdfname.split(" ")) > 1:
                tmp_numpart=tmp_pdfname.split(" ")[0]
                tmp_titlepart="".join(tmp_pdfname.split(" ")[1:])
                tmp_titlepart=re.sub("\W","",tmp_titlepart)
                new_fn="\\\\".join(fn.split("\\\\")[:-1] + [tmp_numpart + " " + tmp_titlepart + ".pdf"])
            else:
                tmp_pdfname=re.sub("\W","",tmp_pdfname)
                new_fn="\\\\".join(fn.split("\\\\")[:-1] + [tmp_pdfname + ".pdf"])

            if new_fn != fn:
                os.rename(fn,new_fn)
                # update fn
                fn=new_fn
                logger.debug(f'pdfname changed to {tmp_pdfname}')
    ########################################################
            fhandle=open(fn, 'rb')
            reader=PdfFileReader(fhandle, strict=False)

            if reader.isEncrypted:
                logger.debug("**********is encrypted pdf**********")
                pdfinputfn=fn
                pdfoutputfn=fn.split(".pdf")[0]+"_resave.pdf"
                if not os.path.isfile(pdfoutputfn):
                    logger.debug("**********resave needed**********")
                    subprocess.run([exe, "-o",pdfoutputfn,"-sDEVICE=pdfwrite", pdfinputfn])
                #replace with the new resaved pdf
                fhandle.close()
                fn=pdfoutputfn
                fhandle = open(pdfoutputfn, 'rb')
                reader=PdfFileReader(fhandle, strict=False)

            for ind in range(reader.numPages):
                page=reader.getPage(ind)
                if round(page.mediaBox.getWidth()) not in [round(pagesizes.A4[0]), round(pagesizes.A4[1])]:
                    logger.debug("**********not A4 size**********")
                    pdfinputfn=fn
                    pdfoutputfn=fn.split(".pdf")[0]+"_fitpage.pdf"
                    if not os.path.isfile(pdfoutputfn):
                        logger.debug("**********fitpage needed**********")
                        subprocess.run([exe, "-o",pdfoutputfn,"-sDEVICE=pdfwrite", "-dPDFFitPage", "-dUseArtBox", pdfinputfn])
                    #replace with the new resaved pdf
                    fhandle.close()
                    fhandle = open(pdfoutputfn, 'rb')
                    reader=PdfFileReader(fhandle, strict=False)
                    break

            bookmarkstr="_".join(fn.strip().split("\\\\")[-2:])

            df.loc[index,"pdf_pagecount"]=reader.numPages
            if index == 0:
                df.loc[index,"start_pagenum"] = 1
            else:
                df.loc[index,"start_pagenum"]=df.loc[:index-1,["pdf_pagecount","pdf_append_blankpage"]].sum().sum()+1

            tmp_start_pagenum=df.loc[index,"start_pagenum"]
            # print(tmp_start_pagenum)
            #确保特定页均为奇数页，以便打印
            prevfn=("" if index==0 else fnlist[index-1])
            if isFixedOddPage(index,fn,prevfn) and (tmp_start_pagenum%2 == 0):
                df.loc[index-1,"pdf_append_blankpage"]=1
                df.loc[index,"start_pagenum"]=df.loc[index,"start_pagenum"]+1
                merger.append(tempPDF)

            merger.append(reader,bookmarkstr)
            # fhandle.close()
        merger.write(outputhandle)
    logger.info("==================== Finished merging pdfs ====================")
    
    orderdf=pd.read_excel(orderfn,sheet_name="Sheet1",usecols=["件号","文号","题名"],dtype={"件号":int})
    orderdf["foldername"]=orderdf["文号"]+" "+orderdf["题名"].apply(lambda x:re.sub("\W","",x))
    orderdf.set_index("foldername",drop=True,inplace=True)
    orderdf.drop(columns=["文号","题名"],inplace=True)
    
    if flag == "发文":        
        for s in set(foldername):
            df.loc[(df.title==s) & (df.pdf=="正文.pdf") , "stamp_pagecount"]=df.groupby("title").pdf_pagecount.sum()[s]
            df.loc[(df.title==s) & (df.pdf=="正文.pdf") , "stamp_order"]=orderdf.loc[s].values[0]
            df.loc[df.title==s , "subtract_pagecount"]=df.loc[(df.title==s) & (df.pdf=="正文.pdf") , "start_pagenum"]-1
    elif flag == "收文":
        foldername_receive=[s for s in foldername if not re.match("\AXX(北方|京)",s)]
        countstartind=0
        for ind, row in df.iterrows():
            if ind>0 and (row.title in foldername_receive) and row.pdf=="文单.pdf":
                df.loc[countstartind, "stamp_pagecount"] = df.iloc[countstartind:ind].pdf_pagecount.sum()
                df.loc[countstartind:ind, "subtract_pagecount"] = df.loc[countstartind, "start_pagenum"]-1
                df.loc[countstartind, "stamp_order"] = orderdf.loc[df.loc[countstartind,"title"]].values[0]
                countstartind=ind
            elif ind==len(df)-1:
                df.loc[countstartind, "stamp_pagecount"] = df.iloc[countstartind:].pdf_pagecount.sum()
                df.loc[countstartind:, "subtract_pagecount"] = df.loc[countstartind, "start_pagenum"]-1  
                df.loc[countstartind, "stamp_order"] = orderdf.loc[df.loc[countstartind,"title"]].values[0]
    logger.info("==================== Finished creating df ====================")
    return df

def replaceWithGrayPagespdf(toconvertfn,inputfn="overlayed.pdf",substitutegraypagefn="convertedgray.pdf",outputfn="toprint.pdf"):
    toSubstitutePages=loadPagenums(toconvertfn)
    fullpdfReader=PdfFileReader(open(inputfn, 'rb'))
    graypdfReader=PdfFileReader(open(substitutegraypagefn,'rb'))
    writer = PdfFileWriter()
    count=0
    for ind in range(fullpdfReader.numPages):
        if ind+1 not in toSubstitutePages:
            writer.addPage(fullpdfReader.getPage(ind))
        else:
            writer.addPage(graypdfReader.getPage(count))
            count=count+1
    with open(outputfn,"wb") as fp:
        writer.write(fp)
    logger.info("==================== Finished replacing with gray pages ====================")

if __name__ == "__main__":
     mergepdf("pdflist_receive_test.txt")
    # replaceWithGrayPagespdf(....)