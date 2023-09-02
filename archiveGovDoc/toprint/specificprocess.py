import os,sys
# from docx2pdf import convert
import pandas as pd
sys.path.append(os.path.abspath(".."))
from preprocess.filter4archive import filter_wb





if __name__ == "__main__":
	# fn=r"C:\Users\Amy19\Documents\0000-文书归档\2021\归档资料—其他\2021年发文登记簿（原北京）.xlsx"
	# fn=r"C:\Users\Amy19\Documents\0000-文书归档\2021\归档资料—其他\2021年发文登记簿.xlsx"
	fn=r"C:\Users\Amy19\Documents\0000-文书归档\2021\归档资料—其他\2021年收文登记簿.xlsx"

	resfn = os.path.splitext(fn)[0] + "_打印用.xlsx"
	filter_wb(fn, resfn)

