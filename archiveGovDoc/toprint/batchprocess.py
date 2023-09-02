# -*- coding: utf-8 -*-

import os
from docx2pdf import convert
import logging
logging.basicConfig(filename='无oa流程.log', encoding='utf-8', level=logging.DEBUG)
import re
from shutil import move
import glob
import win32com.client as win32
from win32com.client import constants

def get_fulldir_list(rootpath,foldernamefn):
    flag=rootpath.split("\\\\")[-2]
    txt = open(foldernamefn, "r",encoding="utf-8")
    l=txt.readline()
    fulldirlist=[]
    while l:
        #ignore blank lines represented by "\n"
        if l.strip():
            tmprootpath=rootpath
            #行首缩进表示待合并请示
            if (re.match(r"^\s",l) and flag=="收文"):
                if re.search("\AXX京",l.strip()):
                    tmprootpath=rootpath.rsplit("收文\\\\",1)[0]+"发文（原北京）\\\\"
                elif re.search("\AXX北方",l.strip()):
                    tmprootpath=rootpath.rsplit("收文\\\\",1)[0]+"发文\\\\"
            fulldir=tmprootpath+l.strip()
            fulldirlist.append(fulldir)
        l=txt.readline()
    txt.close()
    return flag,fulldirlist 

def doc_to_docx(path):
    # Rename path with .docx
    new_file_abs = os.path.abspath(path)
    new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)
    # alternative way
    # new_file_abs = os.path.splitext(os.path.abspath(path))[0] + ".docx"
    if os.path.isfile(new_file_abs):
        # print("exists docx ver:")
        # print("_".join(new_file_abs.split("\\")[-2:]))
        return
    else:       
        # Opening MS Word
        word = win32.gencache.EnsureDispatch('Word.Application')
        doc = word.Documents.Open(path)
        doc.Activate()
    
        # Save and Close
        word.ActiveDocument.SaveAs(
            new_file_abs, FileFormat=constants.wdFormatXMLDocument
        )
        doc.Close(False)
        # print("doc to docx conversion completed:")
        # print("_".join(new_file_abs.split("\\")[-2:]))
        return


def docx_exists(fpath):
    str1=os.path.splitext(fpath)[0]
    str2=".docx"
    return os.path.isfile(str1+str2)


def pdfver_exists(fpath):
    str1=os.path.splitext(fpath)[0]
    str2=".pdf"
    return os.path.isfile(str1+str2)


def get_fname_simple(fname):
    str1=os.path.splitext(fname)[0]
    str2=os.path.splitext(fname)[1]
    if len(str1.split(" "))>1 and (not str1.startswith("附件")):
        # eg: '附件1：债权债务抵销协议书.pdf'
        tmp_numpart=str1.split(" ")[0]
        tmp_titlepart= "".join(str1.split(" ")[1:])      
        tmp_titlepart= re.sub("\W","",tmp_titlepart)
        clean_str1= tmp_numpart + " " + tmp_titlepart
    else:
        clean_str1=re.sub("\W","",str1)
    fname_simple=clean_str1+str2
    return fname_simple
  

def preprocess_dirs(fulldirlist):
    for curfolder in fulldirlist: 
        #curfolder = "C:\\Users\\Amy19\\Documents\\0000-文书归档\\2021\\归档资料—打印文件夹\\发文（原北京）\\XX京安〔2021〕103号 关于表彰北京分公司第二届安全知识竞赛活动获奖团体及个人的决定"
        #simplify filenames and convert doc/docx to pdf per folder
        print("Preprocessing folder " + curfolder.split("\\")[-1])
        #first rename only
        tmpl=list(glob.iglob(curfolder + "\\*.*", recursive=True))
        srcl= [i for i in tmpl if os.path.isfile(i)]
        for f in srcl:
            fname_simple = get_fname_simple(os.path.basename(f))
            fpath = curfolder + r"\\" + fname_simple
            if fname_simple != os.path.basename(f):
                # force overwrite
                move(f, fpath)
                print(f'Renamed from {os.path.basename(f)} to {fname_simple}')
        #then judge and convert
        tmpl=list(glob.iglob(curfolder + "\\*.*", recursive=True))
        srcl= [i for i in tmpl if os.path.isfile(i)]
        for f in srcl:
            if f.endswith(".doc") and (not docx_exists(f)):
                doc_to_docx(f)
            elif f.endswith(".docx") and (not pdfver_exists(f)):
                print("Converting: " + f)
                convert(f)


# require preprocess to simple version fname
def is_needed_pdf(fname):
    if not fname.endswith(".pdf"):
        return False
    elif fname.endswith("_resave.pdf") or fname.endswith("_fitpage.pdf"):
        return False
    else:
        return True


def toPrintPdfList(flag, fulldirlist, toprintpdflistfn):
    with open(toprintpdflistfn,"w",encoding="utf-8") as f:
        for curfolder in fulldirlist:
            tmpl = list(glob.iglob(curfolder + "\\*.*", recursive=True))
            srcl = [i for i in tmpl if os.path.isfile(i)]

            tmp_curfolderfnlist=[]
            tmp_curfolder_dinggaofnlist=[] #for send type only
            
            for j in srcl:
                fname = os.path.basename(j)
                if is_needed_pdf(fname):
                    if re.search("定稿", fname): #多类定稿
                        tmp_curfolder_dinggaofnlist.append(curfolder+"\\\\"+fname+"\n")
                    elif re.search("(正文|文单|处理记录|定稿)",fname) is None: #作为独立文档的管理细则等正文、附件
                            if re.sub("\W","",curfolder.split("\\")[-1]) not in re.sub("\W","",fname): #考虑人资职务任免请示会额外合并pdf的情况
                                tmp_curfolderfnlist.append(curfolder+"\\\\"+fname+"\n")
                            elif (flag=="收文") & ("含附件" not in fname):
                                #若为收文对应的请示发文文件夹，则将盖章后pdf版本作为正文
                                tmp_curfolder_zhengwenfn = curfolder+"\\\\"+fname+"\n"
            
            #若无文单则判断无oa流程，写入log
            fn_wendan=curfolder+"\\\\"+"文单.pdf"
            fn_record=curfolder+"\\\\"+"处理记录.pdf"
            if not os.path.isfile(fn_wendan):
                logging.warning("%s", curfolder.split("\\")[-1])       
            
            #收文对应请示，按旧规定放在复文后，合为一件
            #此处以区分收文待打印文件夹列表，即 foldername_receive.txt 中的请示类发文
            tmp=curfolder.split("\\")[-1]
            flag_send=re.search("\AXX(北方|京)",tmp)

            # 发文（原北京）
            if ("发文" in flag) or flag_send:
                if os.path.isfile(fn_wendan):
                    tmp_curfolderfnlist.append(fn_wendan+"\n")
                    tmp_curfolderfnlist.append(fn_record+"\n") #有文单则可认为同时有处理记录
                tmp_curfolderfnlist.insert(0,curfolder+"\\\\"+"正文.pdf"+"\n") #发文类，最先打印正文
                tmp_curfolderfnlist.extend(tmp_curfolder_dinggaofnlist )#发文类，最后打印定稿
            if flag=="收文":
                if flag_send:
                    #正文替换为盖章后pdf版本
                    tmp_curfolderfnlist[0]=tmp_curfolder_zhengwenfn
                else: 
                    #默认收文类均有文单和处理记录，不作条件判断
                    tmp_curfolderfnlist.insert(0,fn_wendan+"\n") #收文类，最先打印文单
                    tmp_curfolderfnlist.insert(1,tmp_curfolder_zhengwenfn) #收文类，第二个打印正文，只取原始版本，其余版参考 is_needed_pdf(fname)
                    #2022.3.18 updated: 收文不再需要打印处理记录
                    # tmp_curfolderfnlist.append(fn_record+"\n") #收文类，最后打印处理记录
            
            f.writelines(tmp_curfolderfnlist)                



def main(rootpath,foldernamefn,toprintpdflistfn):
    flag, fulldirlist=get_fulldir_list(rootpath, foldernamefn)
    preprocess_dirs(fulldirlist)
    toPrintPdfList(flag, fulldirlist, toprintpdflistfn)
    
################################ for test use ###########################################
# rootpath=r"C:\\Users\\Amy19\\Documents\\documents\\001-python\\文书归档-合并pdf打印用\\minimal source folder\\发文\\"
# foldernamefn="foldername_send_test.txt"
# toprintpdflistfn="pdflist_send_test.txt"
# main(rootpath,foldernamefn,toprintpdflistfn)

# rootpath=r"C:\\Users\\Amy19\\Documents\\documents\\001-python\\文书归档-合并pdf打印用\\minimal source folder\\收文\\"
# foldernamefn="foldername_receive_test.txt"
# toprintpdflistfn="pdflist_receive_test.txt"
# main(rootpath,foldernamefn,toprintpdflistfn)
#########################################################################################

## rootpath=r"C:\\Users\\Amy19\\Documents\\0000-文书归档\\2021\\归档资料—打印文件夹\\发文（原北京）\\"

# rootpath=r"C:\\Users\\Amy19\\Documents\\0000-文书归档\\2021\\归档资料—打印文件夹\\发文\\"
# foldernamefn="foldername_send.txt"
# toprintpdflistfn="pdflist_send.txt"
# main(rootpath,foldernamefn,toprintpdflistfn)
#
# rootpath=r"C:\\Users\\Amy19\\Documents\\0000-文书归档\\2021\\归档资料—打印文件夹\\收文\\"
# foldernamefn="foldername_receive.txt"
# toprintpdflistfn="pdflist_receive.txt"
# main(rootpath,foldernamefn,toprintpdflistfn)



if __name__ == "__main__":
    # rootpath=r"C:\\Users\\Amy19\\Documents\\0000-文书归档\\2021\\归档资料—打印文件夹\\发文（原北京）\\"
    rootpath=r"C:\\Users\\Amy19\\Documents\\0000-文书归档\\2021\\归档资料—打印文件夹\\收文\\"
    foldernamefn="foldername_receive.txt"
    toprintpdflistfn="pdflist_receive.txt"
    main(rootpath,foldernamefn,toprintpdflistfn)