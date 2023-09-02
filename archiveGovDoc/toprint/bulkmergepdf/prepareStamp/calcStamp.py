import os
import re

import logging.config
import yaml

with open('./config.yaml', 'r') as f:
    config = yaml.safe_load(f.read())
    logging.config.dictConfig(config)
logger = logging.getLogger('verboseLogger')


def calcStamp(df):
    # subtractnumList=[]
    stampPcountList=[]
    stampOrderList=[]
    stampPnoList=[]
    blankpagePnumList=[]
    logger.info("==================== Using df to calc stampinfo ====================")
    # only applies to send type
    # for i in range(len(df)):
    #     row=df.iloc[i]
    #     if row.stamp_pagecount !=0:
    #         subtractnum=row.start_pagenum-1
    #         subtractnumList.extend([subtractnum]*row.stamp_pagecount)
    #         stampPcountList=stampPcountList+[row.stamp_pagecount] #store stamp pagecount for cover
    #         stampOrderList=stampOrderList+[row.stamp_order]
    #     elif row.pdf_append_blankpage !=0:
    #         subtractnumList.extend([-1])
    #     else:
    #         continue
    # print(subtractnumList) 
    

    for i in range(len(df)):
        row=df.iloc[i]
        # fullPnoList=fullPnoList+list(range(tmp_start,tmp_end))
        if row.stamp_pagecount >0: # which means cover page
            # subtractnum=row.start_pagenum-1 # initialize for each cover page
            stampPcountList=stampPcountList+[row.stamp_pagecount] #store stamp pagecount for cover
            stampOrderList=stampOrderList+[row.stamp_order]            
            stampPnoList=stampPnoList+list(range(1,row.stamp_pagecount+1))
        # subtractnumList=subtractnumList+[subtractnum]*row.pdf_pagecount
        if row.pdf_append_blankpage == 1:
            # subtractnumList=subtractnumList+[-1]
            blankpagePnumList=blankpagePnumList+[int((row.start_pagenum+row.pdf_pagecount-1)+1)]           
    
    # totalpagenum=df.iloc[-1].start_pagenum+df.iloc[-1].pdf_pagecount-1
    # fullPnoList=list(range(1,int(totalpagenum+1))) # deprecated formula
    # stampPnoList=[(x-delta) if delta>=0 else -1 for (x,delta) in list(zip(fullPnoList,subtractnumList))]
    
    for i in blankpagePnumList:
        stampPnoList.insert(i-1, -999)
 
    logger.debug('**********stampPnoList**********')
    logger.debug(f'**********{stampPnoList}**********')
    logger.debug('**********stampPcountList**********')
    logger.debug(f'**********{stampPcountList}**********')
    logger.debug('**********stampOrderList**********')
    logger.debug(f'**********{stampOrderList}**********') 
    return stampPnoList,stampPcountList,stampOrderList