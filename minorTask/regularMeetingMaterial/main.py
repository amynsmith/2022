from datetime import datetime
import os
import glob as glob

from collectexcel import collectexcel
from splitsheet import splitsheet
from pastetoword import pastetoword

rootfolder = r"C:\Users\Amy19\Documents\科务会材料&总结计划\例会资料收集"
curdirname = "办公例会资料" + datetime.today().strftime("%m%d")
# curdirname = "办公例会资料" + "0801"

wordtemplatefn = os.path.join(rootfolder, "template办公例会资料.docx")

rawexcelfn = os.path.join(rootfolder, "raw" + curdirname + ".xlsx")  # 只分三大版块：总结、计划、待解决问题
excelfn = os.path.join(rootfolder, "res" + curdirname + ".xlsx")  # sheets 最多有10个，包括5个总结、5个计划
wordfn = os.path.join(rootfolder, "res" + curdirname + ".docx")

# 先汇总
rawdir = os.path.join(rootfolder, curdirname)
tmpl = list(glob.iglob(rawdir + "\\*.docx"))
wordfnlist = [fpath for fpath in tmpl if not os.path.basename(fpath).startswith("办公例会资料")]

collectexcel(wordfnlist, rawexcelfn)
splitsheet(rawexcelfn, excelfn)
pastetoword(wordtemplatefn, excelfn, wordfn)

print(f"curdirname: {curdirname}\n")
# input(f"inside {wordfn}, manually add the final section '六、其他工作和需各部门协调解决的相关事宜（5%）'...\n")
