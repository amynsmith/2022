from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
import pandas as pd
import os
import tempfile
import re
import reportlab
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
fontname = "msyh"
pdfmetrics.registerFont(TTFont(fontname,"msyh.ttc"))

import logging
import logging.config
import yaml

with open('./config.yaml', 'r') as f:
    config = yaml.safe_load(f.read())
    logging.config.dictConfig(config)
logger = logging.getLogger('verboseLogger')


# def pnoisUpright(pno, flag):
#     if flag == "发文":
#         # 发文类，页码为奇数时应落在右上角，否则在左上角
#         if pno%2 == 1:
#             return True
#         else:
#             return False
#     else:
#         # flag == "收文"
#         # 收文类，与发文类相反
#         if pno%2 == 0:
#             return True
#         else:
#             return False
#
#
# def generatePnoStampPdf(flag, df,pnostampdir):
#     # max_stamp_pagenum=max(df.groupby("title").pdf_pagecount.sum()) #applies only to send type
#     max_stamp_pagenum=max(df.stamp_pagecount)
#     for pno in range(2,max_stamp_pagenum+1):
#         if os.path.isfile(pnostampdir+str(pno)+".pdf"):
#             continue
#         else:
#             logger.info(f'**********new Pno {pno} Stamp generating**********')
#             c = canvas.Canvas(pnostampdir+str(pno)+".pdf")
#             if pnoisUpright(pno, flag):
#                 c.translate((210)*mm,(297)*mm) #start from topmost upright corner
#                 c.drawString((-10)*mm, (-10)*mm, str(pno))
#             else:
#                 c.translate((0)*mm,(297)*mm) #start from topmost upleft corner
#                 c.drawString((10)*mm, (-10)*mm, str(pno))
#             c.showPage()
#             c.save()


def generatePnoStampPdf(flag, df, pnostampdir):
    # max_stamp_pagenum=max(df.groupby("title").pdf_pagecount.sum()) #applies only to send type
    max_stamp_pagenum=max(df.stamp_pagecount)
    for pno in range(2,max_stamp_pagenum+1):
        if os.path.isfile(pnostampdir+str(pno)+".pdf") and os.path.isfile(pnostampdir + str(pno) + "_alter.pdf"):
            continue
        else:
            logger.info(f'**********new Pno {pno} Stamp generating**********')
            # 2 versions
            if pno%2 == 1:
                c = canvas.Canvas(pnostampdir + str(pno) + ".pdf")
                c.translate((210)*mm,(297)*mm) #start from topmost upright corner
                c.drawString((-10)*mm, (-10)*mm, str(pno))
                c.showPage()
                c.save()
                c = canvas.Canvas(pnostampdir + str(pno) + "_alter.pdf")
                c.translate((0)*mm,(297)*mm) #start from topmost upleft corner
                c.drawString((10)*mm, (-10)*mm, str(pno))
                c.showPage()
                c.save()
            else:
                c = canvas.Canvas(pnostampdir + str(pno) + ".pdf")
                c.translate((0)*mm,(297)*mm) #start from topmost upleft corner
                c.drawString((10)*mm, (-10)*mm, str(pno))
                c.showPage()
                c.save()
                c = canvas.Canvas(pnostampdir + str(pno) + "_alter.pdf")
                c.translate((210)*mm,(297)*mm) #start from topmost upright corner
                c.drawString((-10)*mm, (-10)*mm, str(pno))
                c.showPage()
                c.save()




def generateCoverStampPdf(year,archivenum,stamp_pagecount,coverstampdir):
    coverpdfpath=coverstampdir+str(archivenum)+".pdf"
    if os.path.isfile(coverpdfpath):
        return
    else:
        logger.info(f'**********new Cover Stamp for archivenum {archivenum} generating**********')
        c=canvas.Canvas(coverpdfpath)
        c.setStrokeColorRGB(1, 0, 0)
        c.translate((210)*mm,(297)*mm) #start from topmost upright corner
        leftbottom_x=(-80)*mm
        # leftbottom_y=(-30)*mm
        leftbottom_y=(-25)*mm
        
        innerboxwidth=(20)*mm
        innerboxheight=(10)*mm
        # c.rect((-80)*mm, (-30)*mm, (20)*mm, (10)*mm, stroke=True, fill=False)
        # c.rect((-80)*mm, (-20)*mm, (20)*mm, (10)*mm, stroke=True, fill=False)
        # c.rect((-60)*mm, (-30)*mm, (20)*mm, (10)*mm, stroke=True, fill=False)
        # c.rect((-60)*mm, (-20)*mm, (20)*mm, (10)*mm, stroke=True, fill=False)
        # c.rect((-40)*mm, (-30)*mm, (20)*mm, (10)*mm, stroke=True, fill=False)
        # c.rect((-40)*mm, (-20)*mm, (20)*mm, (10)*mm, stroke=True, fill=False)
        c.rect(leftbottom_x, leftbottom_y, innerboxwidth, innerboxheight, stroke=True, fill=False)
        c.rect(leftbottom_x, leftbottom_y + innerboxheight, innerboxwidth, innerboxheight, stroke=True, fill=False)
        c.rect(leftbottom_x + innerboxwidth, leftbottom_y, innerboxwidth, innerboxheight, stroke=True, fill=False)
        c.rect(leftbottom_x + innerboxwidth, leftbottom_y + innerboxheight, innerboxwidth, innerboxheight, stroke=True, fill=False)
        c.rect(leftbottom_x + 2*innerboxwidth, leftbottom_y, innerboxwidth, innerboxheight, stroke=True, fill=False)
        c.rect(leftbottom_x + 2*innerboxwidth, leftbottom_y + innerboxheight, innerboxwidth, innerboxheight, stroke=True, fill=False)
        
        c.setFont(fontname,10)
        # c.drawCentredString((-70)*mm,(-25)*mm-(5),"XX分公司")
        # c.drawCentredString((-70)*mm,(-15)*mm-(5),"0303")
        # c.drawCentredString((-50)*mm,(-25)*mm-(5),"30年")
        # c.drawCentredString((-50)*mm,(-15)*mm-(5),str(year))
        # c.drawCentredString((-30)*mm,(-25)*mm-(5),str(int(stamp_pagecount))+"页")
        # c.drawCentredString((-30)*mm,(-15)*mm-(5),str(archivenum).zfill(4))
        
        delta=(5) # half of fontsize
        second_row_height=leftbottom_y+innerboxheight/2-delta
        first_row_height=leftbottom_y+innerboxheight+innerboxheight/2-delta
        c.drawCentredString(leftbottom_x + innerboxwidth/2, second_row_height,"XX分公司")
        # c.drawCentredString(leftbottom_x + innerboxwidth/2, second_row_height,"XX分公司")
        c.drawCentredString(leftbottom_x + innerboxwidth/2, first_row_height,"XXXX.0X00")
        c.drawCentredString(leftbottom_x + innerboxwidth + innerboxwidth/2, second_row_height,"10年")
        # c.drawCentredString(leftbottom_x + innerboxwidth + innerboxwidth/2, second_row_height,"10年")
        c.drawCentredString(leftbottom_x + innerboxwidth + innerboxwidth/2, first_row_height, str(year))
        c.drawCentredString(leftbottom_x + 2*innerboxwidth + innerboxwidth/2, second_row_height, str(int(stamp_pagecount))+"页")
        c.drawCentredString(leftbottom_x + 2*innerboxwidth + innerboxwidth/2, first_row_height, str(archivenum).zfill(4))
        
        c.showPage()
        c.save()
