import logging.config
import re
import os
import yaml
import pandas as pd

with open('config.yaml', 'r') as f:
    config = yaml.safe_load(f.read())
    logging.config.dictConfig(config)

import utils.initWebDriver as initdriver
from processProcedure.processall import processall, todo_pageref_dict
from prepareData import bulkdownloadfile
from prepareData import populatedf
from startProcedure.bulkreceivefile import bulkreceivefile
from startProcedure.bulkuploadfile import bulkuploadfile
from prepareData import miscellaneous
from startProcedure.restrictedupload import restrictedupload
from processProcedure.autoprocess import autoprocessbytype


scriptlist = ["processall", "bulkdownloadfile", "populatedf", "bulkreceivefile", "bulkuploadfile", "miscellaneous",
              "restrictedupload", "autoprocess"]


def main(basedir, recfn, target_basedir):
    for ind, l in enumerate(scriptlist):
        print(f"ind {ind}\t {l}")
    tmpstr = input("type ind for script, defaults to 0...\n")
    if tmpstr == "":
        ind = 0
    else:
        ind = [0 if int(tmpstr) not in range(len(scriptlist)) else int(tmpstr)][0]
        if ind == 0:
            driver = initdriver.init()
            while True:
                processall(driver)
                loop = input("continue to process？ else type N or n...\n")
                loop = loop.lower()
                if loop == "n":
                    break
        elif ind == 1:
            resfn = basedir + "\\" + "res_subfoldername.xlsx"
            bulkdownloadfile.main(basedir, resfn)
        elif ind == 2:
            resfn = basedir + "\\" + "respdflist.xlsx"
            resdf = populatedf.populatedf(basedir, resfn)
            populatedf.highlight_respdflist(resfn)
            print("docmark,isattach,recsource,pubdate must be all checked beforehand")
            print("toreceive,toupload are manually added")
            input("After manual check, press to continue...\n")
            resfn = basedir + "\\" + "respdflist_modified.xlsx"
            resdf1 = populatedf.fill_docmark_forattach(resfn)
            resdf2 = populatedf.update_fpath(basedir, resfn)
            # fill recno for all files, including those whose "toreceive" set to False
            resdf3 = populatedf.fill_recno(resfn, recfn)
            input("...\n")
        elif ind == 3:
            driver = initdriver.init()
            fn = basedir + "\\" + "respdflist_modified.xlsx"
            resfn = basedir + "\\" + "res_receivedf.xlsx"
            bulkreceivefile(driver, fn, resfn)
        elif ind == 4:
            driver = initdriver.init()
            # size = driver.get_window_size()
            # driver.set_window_rect(0, 0, size["width"] * 0.5, size["height"])
            fn = basedir + "\\" + "respdflist_modified.xlsx"
            resfn = basedir + "\\" + "res_uploaddf.xlsx"
            bulkuploadfile(driver, fn, resfn)
        elif ind == 5:
            fn = basedir + "\\" + "respdflist_modified.xlsx"
            resdf = miscellaneous.moveto_targetdir(basedir, fn, target_basedir)
            miscellaneous.update_record(fn, recfn)
        elif ind == 6:
            input("edit param first, then uncomment...\n")
            # name1 = "XX XX"
            # name2 = "XX GI G_W UH JW"
            # namestr = name1 + " " + name2
            # branchlist = ["领导班子", None, "XX经理部领导", None]
            # fpath = r"C:\Users\Amy19\Documents\00000-收发文2022\发文\09 纪委发文\3\XXX.pdf"
            # cur_noticetype = 3  # 分公司文件
            # attachpathlist = []
            #
            # cur_title = os.path.splitext(os.path.basename(fpath))[0]
            # inputlist = [cur_title, namestr, branchlist, cur_noticetype]
            # driver = initdriver.init()
            # restrictedupload(driver, inputlist, fpath, attachpathlist)  # revisehref=""
        elif ind == 7:
            processtype = input("choose auto process type...\n")
            if processtype not in todo_pageref_dict.keys():
                print("defaults to auto process 收文 type R...\n")
                processtype = "R"
            # dataprepdir = r"C:\Users\Amy19\PycharmProjects\receiveDoc\prepareData"
            dataprepdir = "prepareData"
            opinionfn = os.path.join(dataprepdir, "records_opinion.xlsx")  # multi sheets corresponding to procedures
            nibanfn = os.path.join(dataprepdir, "extract_niban.xlsx")  # single sheets with multi rows each corresponding to a procedure

            tmpstr = input("data preparation is needed for newly arrived todos, to skip, press S/s...\n")
            prepskip = tmpstr.strip().lower()
            toprep = [False if prepskip=="s" else True][0]
            times_zoomout = 4
            autoprocessbytype(processtype, opinionfn, nibanfn, toprep, times_zoomout)


if __name__ == '__main__':
    logger = logging.getLogger('verboseLogger')
    logging.getLogger("pdfminer").setLevel(logging.WARNING)

    base_path = r"C:\Users\Amy19\Desktop\收文"
    basedir = base_path + pd.Timestamp.today().strftime("%Y%m%d")
    # 严格按照时间顺序，前一天的收文结束并且更新登记簿后，再在此基础上生成内部收文文号
    # basedir = r"C:\Users\Amy19\Desktop\收文20220823"

    target_basedir = r"C:\Users\Amy19\Documents\00000-收发文2022\收文"
    recfn = r"C:\Users\Amy19\Documents\00000-收发文2022\收文\2022年收文登记簿.xlsx"

    main(basedir, recfn, target_basedir)
