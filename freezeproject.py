import glob,os
import py2exe
import re


omitptn = r"C:\\Users\\Amy19\\PycharmProjects\\receiveDoc\\(test|venv)"

rootpath = r"C:\Users\Amy19\PycharmProjects\receiveDoc"
pyfiles = list(glob.iglob(rootpath+"\\**\\*.py", recursive=True))
pyfnlist = []
for fn in pyfiles:
    if re.search(omitptn,fn) is not None:
        continue
    else:
        # relativefn = ".."+fn.split(rootpath,1)[-1]
        pyfnlist.append(fn)

# print(pyfnlist)

# https://github.com/moses-palmer/pynput/issues/312#issuecomment-710110753
includes = ["pynput.keyboard._win32", "pynput.mouse._win32"]

py2exe.freeze(
    console=pyfnlist,
    data_files=[("prepareData", ["C:/Users/Amy19/PycharmProjects/receiveDoc/prepareData/extract_niban.xlsx", "C:/Users/Amy19/PycharmProjects/receiveDoc/prepareData/records_opinion.xlsx"]),
                ("resources", ["C:/Users/Amy19/PycharmProjects/receiveDoc/resources/icon_pdfuploadfinish.png", "C:/Users/Amy19/PycharmProjects/receiveDoc/resources/icon_upload.png", "C:/Users/Amy19/PycharmProjects/receiveDoc/resources/icon_open.png"]),
                (".", ["C:/Users/Amy19/PycharmProjects/receiveDoc/config.yaml"]),
                (".", ["C:/Users/Amy19/PycharmProjects/receiveDoc/utils/chromedriver.exe"])],
    zipfile="library.zip",
    options={
        "py2exe": {
            "includes": includes,
            "bundle_files": 3
        }
    }
)