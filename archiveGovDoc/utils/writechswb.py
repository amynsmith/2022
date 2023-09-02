import pandas as pd

def str_len(str):
    try:
        row_l=len(str)
        utf8_l=len(str.encode('utf-8'))
        return (utf8_l-row_l)/2+row_l
    except:
        return None


def writechswb(resdf,resfn):
    Excelwriter = pd.ExcelWriter(resfn,engine="xlsxwriter")
    resdf.to_excel(Excelwriter,sheet_name="0",index=False)
    worksheet = Excelwriter.sheets["0"]  # pull worksheet object
    for idx, col in enumerate(resdf):  # loop through all columns
        series = resdf[col]
        max_len = max((
            series.astype(str).map(str_len).max(),  # len of largest item
           str_len(str(series.name))  # len of column name/header
            )) + 2  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width
    Excelwriter.save()