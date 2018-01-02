# -*- coding: utf-8 -*-
import xlrd,os

def readexcel_data(sheet_value,row_value,col_value):
    DIR = os.path.dirname(os.path.dirname(__file__))
    filename = os.path.join(DIR,'test_data','interface.xlsx')
    data = xlrd.open_workbook(filename)
    sheet = data.sheet_by_index(sheet_value)
    datarow = sheet.row(row_value)[col_value].value
    return datarow