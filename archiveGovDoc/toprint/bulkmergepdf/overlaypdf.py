from prepareStamp.calcStamp import calcStamp
from prepareStamp.generateStamp import generatePnoStampPdf,generateCoverStampPdf

from PyPDF2 import PdfFileReader, PdfFileWriter
import re

import logging.config
import yaml

with open('config.yaml', 'r') as f:
    config = yaml.safe_load(f.read())
    logging.config.dictConfig(config)
logger = logging.getLogger('verboseLogger')


def overlaypdf(flag,df,year,pnostampdir,coverstampdir,mergedfn="merged_original.pdf",finalresfn="finalres.pdf"):
    generatePnoStampPdf(flag, df,pnostampdir) # generate stamp pdfs
    stampPnoList,stampPcountList,stampOrderList=calcStamp(df)
    reader=PdfFileReader(open(mergedfn, 'rb'))
    writer = PdfFileWriter()
    totalpagenum=reader.numPages
    coverind=0
    logger.info("==================== Start overlaying pdfs ====================")
    for ind in range(totalpagenum):
        page=reader.getPage(ind)
        box = page.mediaBox
        deg = page.get('/Rotate')
        # print(f"deg: {deg}")
        # print(f"box: {box}")
        logger.info(f"page {ind + 1} deg: {deg}")

        stampPageNum=stampPnoList[ind]
        if stampPageNum<0: #corresponding to blank page
            pass
        elif stampPageNum != 1:
            # 每件的扉页不加页码，其余非空页加页码
            if (ind + 1)% 2 == (stampPageNum%2):
                # 实际打印用页码和件内页码，奇偶性一致
                stampfn=pnostampdir+str(int(stampPageNum))+".pdf"
            else:
                logger.info(f"using {stampPageNum}_alter.pdf")
                stampfn = pnostampdir + str(int(stampPageNum)) + "_alter.pdf"

            stampPdf=PdfFileReader(open(stampfn, "rb"))
            stampPage=stampPdf.getPage(0)

            if box.upperRight[0] > box.upperRight[1]:
                if deg not in [0, 180, None]:
                    ######## deg=90 deg=270 tested
                    page.mergeRotatedTranslatedPage(stampPage, deg, stampPage.mediaBox.getWidth()/2, stampPage.mediaBox.getWidth()/2)

                else:
                    # if deg in [0, 180, None]:
                    # 原页面为横向布局，叠加纵向布局的页码pdf
                    # both use counterClockWise
                    page.rotateCounterClockwise(90).mergeRotatedTranslatedPage(stampPage, -90,
                                                                               stampPage.mediaBox.getWidth() / 2,
                                                                               stampPage.mediaBox.getWidth() / 2)
                    print(f'page {ind + 1} is rotated to portrait')
            else:
                page.mergePage(stampPage)
        else:
            # 一件中的扉页，加盖归档章
            logger.debug(f'page {ind+1} is coverpage')
            #add more complex stamp for coverpage            
            stamp_pagecount=stampPcountList[coverind]
            archivenum=stampOrderList[coverind]
            coverind=coverind+1
            generateCoverStampPdf(year,archivenum,stamp_pagecount,coverstampdir)
            coverpdfpath=coverstampdir+str(archivenum)+".pdf"
            page.mergePage(PdfFileReader(open(coverpdfpath,"rb")).getPage(0))
        writer.addPage(page)

    with open(finalresfn,"wb") as fp:
        writer.write(fp)   
    logger.info("==================== Finished overlaying pdfs ====================")