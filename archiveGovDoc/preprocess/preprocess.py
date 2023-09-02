import pandas as pd
import os
from filter4archive import applyfilter
from tag import addtag
from sort import customsort
from format import addformat


def main(fn_send, fn_receive, resfolder):
    tmpresfn = os.path.join(resfolder, "emptytag.xlsx")
    applyfilter(fn_send, fn_receive, tmpresfn)

    fn=tmpresfn
    tmpresfn = os.path.join(resfolder, "addtag.xlsx")
    addtag(fn, tmpresfn)

    fn=tmpresfn
    tmpresfn = os.path.join(resfolder, "res_sort.xlsx")
    customsort(fn, tmpresfn)

    fn=tmpresfn
    tmpresfn = os.path.join(resfolder, "res_format.xlsx")
    addformat(fn, tmpresfn)

    return




if __name__ == "__main__":
    # fn_send="raw//2021年发文登记簿（原北京）.xlsx"
    fn_send="raw//2021年发文登记簿.xlsx"

    # fn_send = "raw//2021年发文登记簿.xlsx"
    fn_receive = "raw//2021年收文登记簿.xlsx"


    resfolder="sort"
    main(fn_send, fn_receive, resfolder)