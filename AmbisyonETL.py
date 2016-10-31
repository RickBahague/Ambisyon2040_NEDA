#Ambisyon2040 Survey Data Set XLSX/XLS to CSV parser
#rick (at) cp-union (dot) com

import pandas as pd
import openpyxl
import os

def get_data(fname,data_path):
    wb = openpyxl.load_workbook(data_path + fname)
    sheet = wb.get_sheet_by_name(wb.sheetnames[0])
    cols = len(sheet.columns)
    rows = len(sheet.rows)
    lines = dict()
    for r in range(0, rows):
        d = dict()
        for c in range(0,cols):
            d[c] = sheet.rows[r][c].value
        lines[r] = d
    df=pd.DataFrame(lines).T
    df['fname'] = fname[:-5]
    question = df.ix[3][0]
    df['question'] = question
    data = df[(pd.isnull(df[0])==False) | (pd.isnull(df[1])==False)]
    cols_response = df.ix[8].values[0:]
    cols_response[0] = 'Response'
    data.columns = cols_response
    data.reset_index(inplace=True)
    data = data.drop(["index"],axis=1)
    data = data[7:][data.columns[:-2]]
    data['question'] = question
    return data
