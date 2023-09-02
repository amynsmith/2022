#!/usr/bin/env python
# coding: utf-8
import shutil
import xml.etree.ElementTree as ET
import os
import pandas as pd
from shutil import make_archive


def getxmlstructure(xmlfn):
    tree = ET.parse(xmlfn)
    root = tree.getroot()
    # root.tag == "data"
    # print(root.attrib)
    tmp_direc_desc = root.findall("./Package/Description")[0]
    tmp_file_desc = root.findall("./Package/Content/Package/Description")[0]
    tags_directory = []
    for i in tmp_direc_desc.findall("./*"):
        tags_directory = tags_directory + [i.tag]
    tags_file = []
    for i in tmp_file_desc.findall("./*"):
        tags_file = tags_file + [i.tag]
    return tags_directory, tags_file



def buildxml(xlsfn_direc, xlsfn_file, xmlfn, resfn):
    tags_directory, tags_file = getxmlstructure(xmlfn)
    df_direc = pd.read_excel(xlsfn_direc, dtype=str, keep_default_na=False)
    df_file = pd.read_excel(xlsfn_file, dtype=str, keep_default_na=False)
    # createdate="2022-4-28 9:10:43"
    createdate = pd.Timestamp.today().strftime("%F %T")
    data = ET.Element("data")
    data.set("struLevelAttribute", "2")
    data.set("createdate", createdate)
    data.set("version", "1.0")
    # add filter column
    df_file["directoryid"] = 0
    directoryid = 0
    for ind, row in df_file.iterrows():
        if ind == 0:
            df_file.loc[ind, "directoryid"] = 0
        elif row["文件类别"] != "正文":
            df_file.loc[ind, "directoryid"] = df_file.iloc[ind - 1].directoryid
        elif row["文件类别"] == "正文":
            directoryid = directoryid + 1
            df_file.loc[ind, "directoryid"] = directoryid
    # per folder
    for ind, row in df_direc.iterrows():
        Package = ET.SubElement(data, "Package")
        Description = ET.SubElement(Package, "Description")
        for tag in tags_directory:
            new_element = ET.SubElement(Description, tag)
            new_element.text = df_direc.iloc[ind][tag]
            # print(new_element.tag, new_element.text)
        Content = ET.SubElement(Package, "Content")
        tmpdf = df_file[df_file.directoryid == ind]
        for subind, row in tmpdf.iterrows():
            Package = ET.SubElement(Content, "Package")
            Description = ET.SubElement(Package, "Description")
            for tag in tags_file:
                new_element = ET.SubElement(Description, tag)
                new_element.text = tmpdf.loc[subind, tag]
                # print(new_element.tag, new_element.text)
    # ET.dump(data)
    tree = ET.ElementTree(element=data)
    tree.write(resfn, encoding="utf-8", xml_declaration=True)
    return


def main(basedir, direcxls, filexls, templatexml, resxml, zipdir):
    xlsfn_direc = os.path.join(basedir, direcxls)
    xlsfn_file = os.path.join(basedir, filexls)
    xmlfn = os.path.join(basedir, templatexml)
    resfn = os.path.join(basedir, resxml)
    buildxml(xlsfn_direc, xlsfn_file, xmlfn, resfn)
    dstfn=os.path.join(zipdir,resxml)
    shutil.copy(resfn, dstfn)
    basename = os.path.dirname(zipdir)+"\\"+"uploadtest"
    make_archive(basename, "zip", root_dir=zipdir, base_dir=".")





if __name__ == "__main__":
    basedir = r"C:\Users\Amy19\Desktop\xml"
    direcxls = "企业管理类卷内级.xls"
    filexls = "企业管理类电子文件级.xls"
    templatexml = "sip-data.xml"
    resxml = "resupload.xml"
    zipdir = r"C:\Users\Amy19\Desktop\uploadtest xml"
    main(basedir, direcxls, filexls, templatexml, resxml, zipdir)





