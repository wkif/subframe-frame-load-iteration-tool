import csv
import codecs
from pandas.io.excel import ExcelWriter
import pandas as pd
import xlsxwriter
import xlrd as xlrd
import os
import re
import win32file

def CsvToJson(csvPath):
    with open(csvPath, "r") as csvfile:
        reader = csv.reader(csvfile)
        rows = [row for row in reader]
        blockNmaeList = []
        for j in range(1, len(rows[0])):
            if (rows[0][j] and rows[0][j] != ' '):
                blockNmaeList.append(rows[0][j])
        return blockNmaeList

class Data():
    le = 0
    ri = 0
def getxlsxData(path, blockName):
    list = path.rsplit("/", 1)
    filename = list[1]
    filepath = list[0]

    xlsxname = filename.split('.')[0] + '.xlsx'
    xlsxPath = os.path.join(filepath, xlsxname)
    workbook = xlsxwriter.Workbook(xlsxPath)
    worksheet = workbook.add_worksheet('first_sheet')
    workbook.close()
    with ExcelWriter(xlsxPath) as ew:
        pd.read_csv(path).to_excel(ew, index=False)
    demoxlsxpath = readData(filepath, xlsxname, blockName)

    ans = [Data.le, Data.ri, demoxlsxpath]
    return ans


def readData(filepath, xlsxname, blockName):
    xlsxpath = os.path.join(filepath, xlsxname)
    # 打开文件
    workbook = xlrd.open_workbook(xlsxpath)
    # 根据sheet索引或者名称获取sheet内容
    sheet = workbook.sheet_by_index(0)  # sheet索引从0开始

    rows = sheet.nrows  # 获取有多少行
    cols = sheet.ncols  # 获取有多少列
    blocknAME = blockName
    # sheet.cell_value(第几行,第几列)
    demoxlsxpath = os.path.join(filepath, 'cache.xlsx')
    workbook = xlsxwriter.Workbook(demoxlsxpath)  # 创建一个excel文件
    # 在文件中创建一个名为TEST的sheet,不加名字默认为sheet1
    worksheet = workbook.add_worksheet(u'sheet1')

    for i in range(0, rows):
        for j in range(0, 4):
            value = sheet.cell_value(i, j)
            if (re.search('Unnamed', value)):
                value = ''
            worksheet.write(i, j, value)

    left = 0
    right = 0
    flag = 0
    for l in range(4, cols):
        if (flag == 1):
            break
        newvalue = sheet.cell_value(0, l)
        if (newvalue == blocknAME):
            left = l

            flag = 1
            for r in range(l + 1, cols):
                value = sheet.cell_value(0, r)
                searchObj = re.search('Unnamed', value, re.M | re.I)
                if (not searchObj):
                    right = r
                    break
                right = cols
    Data.le = left
    Data.ri = right
    for i in range(0, rows):
        x = 4
        for j in range(left, right):
            value = sheet.cell_value(i, j)
            if (re.search('Unnamed', value)):
                value = ''
            worksheet.write(i, x, value)
            x += 1
    workbook.close()
    return demoxlsxpath


def csv_to_xlsx_pd(cacheFilepath, filepath, filename):
    csvPath = os.path.join(filepath, filename)
    xlsxname = filename.split('.')[0] + '.xlsx'
    xlsxPath = os.path.join(cacheFilepath, 'cache', xlsxname)
    workbook = xlsxwriter.Workbook(xlsxPath)
    worksheet = workbook.add_worksheet('first_sheet')
    workbook.close()
    with ExcelWriter(xlsxPath) as ew:
        pd.read_csv(csvPath).to_excel(ew, index=False)

    return xlsxname


def csv_to_xlsx(filepath, filename):
    csvPath = os.path.join(filepath, filename)
    xlsxname = filename.split('.')[0] + '.xlsx'
    xlsxPath = os.path.join(filepath, xlsxname)
    workbook = xlsxwriter.Workbook(xlsxPath)  # 创建一个excel文件
    worksheet = workbook.add_worksheet(u'Sheet1')  # 在文件中创建一个名为TEST的sheet,不加名字默认为sheet1
    with open(csvPath, 'r', encoding='utf-8') as f:
        read = csv.reader(f)
        l = 0
        for line in read:
            r = 0
            for i in line:
                if (re.search('noname', i)):
                    pass
                elif (re.search(' ', i)):
                    pass
                else:
                    worksheet.write(l, r, i)
                r = r + 1
            l = l + 1
    workbook.close()
    return xlsxname


def xlsx_to_csv(xlsxpath):
    list = xlsxpath.rsplit("/", 1)
    filename = list[1]
    filepath = list[0]
    csvname = filename.split('.')[0] + '.csv'
    csvPath = os.path.join(filepath, csvname)
    workbook = xlrd.open_workbook(xlsxpath)
    table = workbook.sheet_by_index(0)
    with codecs.open(csvPath, 'w', encoding='utf-8') as f:
        write = csv.writer(f)
        for row_num in range(table.nrows):
            row_value = table.row_values(row_num)
            # print(row_value)
            for i in range(len(row_value)):
                if (re.search('Unnamed', row_value[i])):
                    row_value[i] = ''
            write.writerow(row_value)

def is_used(file_name):
    try:
        v_handle = win32file.CreateFile(file_name, win32file.GENERIC_READ, 0, None, win32file.OPEN_EXISTING,
                                        win32file.FILE_ATTRIBUTE_NORMAL, None)
        result = bool(int(v_handle) == win32file.INVALID_HANDLE_VALUE)
        win32file.CloseHandle(v_handle)
    except Exception:
        return True
    return result
