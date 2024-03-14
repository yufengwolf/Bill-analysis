import xlrd,xlwt
from xlutils.copy import copy
import datetime
import os
import pdfplumber
from PIL import Image
from pypdf import PdfMerger as pm

# 读取指定文件后缀名文件
def filter_file_name(name, file_type):
    path = os.getcwd()
    list_dir = os.listdir(path)
    excel_names = []
    for i in list_dir:
        s = i.split(".")
        if s[-1] == file_type and name in s[0]:
            excel_names.append(i)
    return excel_names

#读取储蓄卡账单明细
def cash_card_input(worksheet, file_name):
    file_name = filter_file_name(file_name, 'xls')
    print(file_name)
    for name in file_name:
        workbook = xlrd.open_workbook(name,formatting_info=True)
        sheets = workbook.sheet_names()
        new_row_index = 0
        new_col_index = 0
        for sheetIndex in sheets:
            insert_flag = False
            sheet = workbook.sheet_by_name(sheetIndex)
            for i in range(sheet.nrows):
                if sheet.cell_value(i, 0) == '交易日期' or sheet.cell_value(i, 0) == '交易⽇期':
                    insert_flag = True
                    if new_row_index != 0:
                        continue
                elif '2023' not in str(sheet.cell_value(i, 0)):
                    insert_flag = False
                for j in range(sheet.ncols):
                    if not insert_flag:
                        break
                    worksheet.write(new_row_index, new_col_index, sheet.cell_value(i, j))
                    new_col_index += 1
                if insert_flag:
                    new_row_index += 1
                    new_col_index = 0

#读取信用卡账单明细
def credit_card_input(worksheet, file_name):
    file_name = filter_file_name(file_name, 'xls')
    print(file_name)
    for name in file_name:
        workbook = xlrd.open_workbook(name,formatting_info=True)
        sheets = workbook.sheet_names()
        new_row_index = 0
        new_col_index = 0
        for sheetIndex in sheets:
            sheet = workbook.sheet_by_name(sheetIndex)
            for i in range(sheet.nrows):
                insert_flag = True
                for j in range(sheet.ncols):
                    if sheet.cell_value(i, 0) == '交易日期' and new_row_index == 0:
                        worksheet.write(new_row_index, new_col_index, sheet.cell_value(i, j))
                        new_col_index += 1
                    elif ('2023' not in str(sheet.cell_value(i, 0))) or ('2023' not in str(sheet.cell_value(i, 1))) or ('还款' in str(sheet.cell_value(i, 2))):
                        insert_flag = False
                        break
                    else:
                        worksheet.write(new_row_index, new_col_index, sheet.cell_value(i, j))
                        new_col_index += 1
                if insert_flag:
                    new_row_index += 1
                    new_col_index = 0

#读取账单明细，汇总到excel
def addBandtoExcelTable():
    workbook = xlwt.Workbook(encoding='utf8')
    gongshangSheet = workbook.add_sheet('工商银行')
    nongyeSheet = workbook.add_sheet('农业银⾏')
    jiaotongSheet = workbook.add_sheet('交通银⾏')
    cash_card_input(gongshangSheet, '工商银行')
    cash_card_input(nongyeSheet, '农业银⾏')
    credit_card_input(jiaotongSheet, '交通银行')
    workbook.save('output.xls')

addBandtoExcelTable()
