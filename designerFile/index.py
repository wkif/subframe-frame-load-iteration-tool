# -*- coding: utf-8 -*-
# author:kif<kif101001000@163.com>
# time: 2022,04,29
# Form implementation generated from reading ui file 'connect_me.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!
# 导入程序运行必须模块
import datetime
import os
import re
import shutil
import sys
from openpyxl import load_workbook
import xlrd
from PyQt5 import QtWidgets
from PyQt5.QtCore import QDir, Qt
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QIcon
from openpyxl.styles import PatternFill
from qt_material import apply_stylesheet
import pandas as pd
import numpy as np
from designerFile.tools.tools import getxlsxData, csv_to_xlsx_pd, csv_to_xlsx, xlsx_to_csv, CsvToJson, is_used
from designerFile.mainView import Ui_MainWindow
from designerFile.view2 import Ui_Dialog


class ChildWin_block(QtWidgets.QDialog, Ui_Dialog):
    def __init__(self):
        super(ChildWin_block, self).__init__()
        self.setupUi(self)


class globalData():
    HotSpots_N_List = []
    bolckNameList = []
    file2path = ''
    file1path = ''
    xlsxFile2path = ''
    userFilepath = ''
    BlockCount = 0
    firstFlag = 0
    left = 0
    right = 0
    editChangeFlag = 0
    cacheFilepath = ''
    GVWflag = 0,
    ratioList: []


class MyMainForm(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyMainForm, self).__init__(parent)
        self.setupUi(self)
        # 按钮事件绑定
        self.setWindowIcon(QIcon('F:/File_my/Project/Plug-in_1/file/logo.png'))
        self.pushButton.clicked.connect(self.ImportVRLDA)
        self.pushButton_2.clicked.connect(self.ImportLoadData)
        self.pushButton_3.clicked.connect(self.ImportCYCLE)
        self.pushButton_4.clicked.connect(self.noticeSave)
        self.pushButton_5.clicked.connect(self.writetoCSV)
        self.tableWidget_2.horizontalHeader().sectionClicked.connect(self.HorSectionClicked)  # 表头单击信号

        # tableWidget
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # 左表格铺满整个QTableWidget控件
        self.tableWidget.verticalHeader().setVisible(False)
        # 隐藏列标题
        self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        # 中表格不可编辑
        # font = self.tableWidget.horizontalHeader().font()
        # font.setBold(True)
        # self.tableWidget.horizontalHeader().setFont(font)
        # 标题字体加粗

        # tableWidget_2
        self.tableWidget_2.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tableWidget_2.horizontalHeader().setStyleSheet(
            "QHeaderView::section{font:7pt;color: black;};")

        # 中表格铺满整个QTableWidget控件
        self.tableWidget_2.setEditTriggers(QAbstractItemView.NoEditTriggers)
        # 中表格不可编辑
        self.tableWidget_2.verticalHeader().setVisible(False)

        # tableWidget_5
        self.tableWidget_5.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # 中表格铺满整个QTableWidget控件
        self.tableWidget_5.setEditTriggers(QAbstractItemView.NoEditTriggers)
        # 中表格不可编辑
        self.tableWidget_5.verticalHeader().setVisible(False)  # 隐藏垂直表头
        self.tableWidget_5.horizontalHeader().setVisible(False)  # 隐藏水平表头

        # tableWidget_3
        self.tableWidget_4.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tableWidget_3.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # 右表格铺满整个QTableWidget控件
        self.tableWidget_3.setEditTriggers(QAbstractItemView.NoEditTriggers)
        # 右表格不可编辑
        self.tableWidget_3.verticalHeader().setVisible(False)
        self.tableWidget_4.verticalHeader().setVisible(False)  # 隐藏垂直表头
        self.tableWidget_4.horizontalHeader().setVisible(False)  # 隐藏水平表头
        self.tableWidget_7.verticalHeader().setVisible(False)  # 隐藏垂直表头
        self.tableWidget_7.horizontalHeader().setVisible(False)  # 隐藏水平表头
        self.tableWidget_6.verticalHeader().setVisible(False)  # 隐藏垂直表头
        self.tableWidget_6.horizontalHeader().setVisible(False)  # 隐藏水平表头
        # self.tableWidget_8.verticalHeader().setVisible(False)  # 隐藏垂直表头
        self.tableWidget_8.horizontalHeader().setVisible(False)  # 隐藏水平表头
        self.tableWidget_block = ''

    def ImportVRLDA(self):
        # 实例化QFileDialog
        dig = QFileDialog()
        # 设置可以打开任何文件
        dig.setFileMode(QFileDialog.AnyFile)
        # 文件过滤
        dig.setFilter(QDir.Files)

        if dig.exec_():
            # 接受选中文件的路径，默认为列表
            filenames = dig.selectedFiles()
            # 列表中的第一个元素即是文件路径，以只读的方式打开文件
            if is_used(filenames[0]):
                self.messageDialog('警告', '文件正在被占用')
            else:
                filename = filenames[0].rsplit("/", 1)[1]
                if (not re.search('.csv', filename)):
                    msg_box = QMessageBox(QMessageBox.Warning, '警告', '请选择csv文件')
                    msg_box.exec_()
                else:
                    globalData.file1path = filenames[0]
                    data = pd.read_csv(filenames[0])
                    data2 = np.array(data)
                    data = np.array(data)[::-1]
                    flag = 0
                    self.tableWidget.setRowCount(0)
                    self.tableWidget.clearContents()
                    self.tableWidget_2.setRowCount(0)
                    self.tableWidget_2.clearContents()
                    self.tableWidget_5.setRowCount(0)
                    self.tableWidget_5.clearContents()
                    self.tableWidget_4.setRowCount(0)
                    self.tableWidget_4.clearContents()
                    self.tableWidget_6.setRowCount(0)
                    self.tableWidget_6.clearContents()
                    self.tableWidget_7.setRowCount(0)
                    self.tableWidget_7.clearContents()
                    self.tableWidget_8.setRowCount(0)
                    self.tableWidget_8.clearContents()
                    globalData.HotSpots_N_List.clear()
                    for i in range(len(data)):
                        item = data[i]
                        if (len(item) != 3 and flag == 0):
                            flag = 1
                            self.messageDialog('警告', '文件似乎不是VRLDA损伤数据，确实导入？')

                        globalData.HotSpots_N_List.append(data2[i][1])
                        row = 0
                        self.tableWidget.insertRow(row)
                        for j in range(len(item)):
                            item = QTableWidgetItem(str(data[i][j]))
                            item.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                            self.tableWidget.setItem(row, j, item)

    def ImportLoadData(self):
        dig = QFileDialog()
        dig.setFileMode(QFileDialog.AnyFile)
        dig.setFilter(QDir.Files)
        if dig.exec_():
            filenames = dig.selectedFiles()

            list = filenames[0].rsplit("/", 1)
            filename = list[1]
            if (not re.search('.csv', filename)):
                msg_box = QMessageBox(QMessageBox.Warning, '警告', '请选择csv文件')
                msg_box.exec_()
            elif is_used(filenames[0]):
                self.messageDialog('警告', '文件正在被占用')
            else:
                filepath = list[0]
                globalData.userFilepath = filepath
                datetime_object = datetime.datetime.now()
                time = str(datetime_object).split(' ')[0]
                newFileName = (filename.split('.')[0] + '-' + time + '.csv').replace('-', '_')
                newFilePath = os.path.join(globalData.cacheFilepath, 'cache', newFileName).replace('\\', '/')
                if not os.path.exists(os.path.join(globalData.cacheFilepath, 'cache')):
                    os.makedirs(os.path.join(globalData.cacheFilepath, 'cache'))
                shutil.copy(filenames[0], newFilePath)
                globalData.file2path = newFilePath
                BlockNameList = CsvToJson(globalData.file2path)
                xlsxname = csv_to_xlsx(os.path.join(globalData.cacheFilepath, 'cache'), newFileName)
                globalData.xlsxFile2path = os.path.join(globalData.cacheFilepath, 'cache', xlsxname).replace("\\", '/')
                globalData.bolckNameList = BlockNameList
                globalData.BlockCount = len(BlockNameList)
                self.tableWidget_2.setRowCount(0)
                self.tableWidget_2.clearContents()
                self.tableWidget_5.setRowCount(0)
                self.tableWidget_5.clearContents()
                self.tableWidget_7.setRowCount(0)
                self.tableWidget_7.clearContents()

                self.tableWidget_2.setColumnCount(globalData.BlockCount)
                self.tableWidget_4.setColumnCount(3)
                self.tableWidget_5.setColumnCount(globalData.BlockCount)
                self.tableWidget_7.setColumnCount(globalData.BlockCount)
                ShowbolckNameList = []
                for i in globalData.bolckNameList:
                    ShowbolckNameList.append(i.split(' ')[0])
                self.tableWidget_2.setHorizontalHeaderLabels(ShowbolckNameList)
                self.tableWidget_2.resizeColumnsToContents()

    def ImportCYCLE(self):
        # 实例化QFileDialog
        dig = QFileDialog()
        # 设置可以打开任何文件
        dig.setFileMode(QFileDialog.AnyFile)
        # 文件过滤
        dig.setFilter(QDir.Files)
        if dig.exec_():
            # 接受选中文件的路径，默认为列表
            filenames = dig.selectedFiles()
            list = filenames[0].rsplit("/", 1)
            filename = list[1]
            if (not re.search('.csv', filename)):
                msg_box = QMessageBox(QMessageBox.Warning, '警告', '请选择csv文件')
                msg_box.exec_()
            elif is_used(filenames[0]):
                self.messageDialog('警告', '文件正在被占用')

            else:
                filepath = list[0]
                xlsxname = csv_to_xlsx_pd(globalData.cacheFilepath, filepath, filename)
                # csv转为xlsx
                xlsxPath = os.path.join(globalData.cacheFilepath, 'cache', xlsxname).replace('\\', '/')
                workbook = xlrd.open_workbook(xlsxPath)
                sheet = workbook.sheet_by_index(0)
                rows = sheet.nrows
                cols = sheet.ncols
                self.tableWidget_2.setRowCount(0)
                self.tableWidget_2.clearContents()
                self.tableWidget_4.setRowCount(0)
                self.tableWidget_4.clearContents()
                self.tableWidget_5.setRowCount(0)
                self.tableWidget_5.clearContents()
                self.tableWidget_6.setRowCount(0)
                self.tableWidget_6.clearContents()
                self.tableWidget_7.setRowCount(0)
                self.tableWidget_7.clearContents()
                self.tableWidget_8.setRowCount(0)
                self.tableWidget_8.clearContents()
                self.tableWidget_8.setColumnCount(1)
                self.tableWidget_8.setVerticalHeaderLabels(['TOTAL = '])
                self.tableWidget_8.setEditTriggers(QAbstractItemView.NoEditTriggers)
                self.tableWidget_7.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
                self.tableWidget_6.setColumnCount(1)
                self.tableWidget_6.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
                self.tableWidget_6.setEditTriggers(QAbstractItemView.NoEditTriggers)
                dataListfortableWidget_2 = []
                zeroList = []
                dataListfortableWidget_4 = []
                dataListfortableWidget_5 = []
                for i in range(len(globalData.bolckNameList) + 1):
                    zeroList.append(0)
                for item in globalData.HotSpots_N_List:
                    flag = 0
                    for r in range(1, rows):
                        if (item == sheet.cell_value(r, 0)):
                            ltemList = []
                            for block in globalData.bolckNameList:
                                flag2 = 0
                                for l in range(1, cols):
                                    s = sheet.cell_value(0, l)
                                    if (re.search(s, block)):
                                        ltemList.append(sheet.cell_value(r, l))
                                        flag2 = 1
                                        break
                                if (flag2 == 0):
                                    ltemList.append(0)
                            dataListfortableWidget_2.append(ltemList)
                            flag = 1
                            break
                    if (flag == 0):
                        dataListfortableWidget_2.append(zeroList)
                        dataListfortableWidget_4.append(
                            ['ADD', sheet.cell_value(globalData.HotSpots_N_List.index(item) + 1, 0)])
                        ltemList = []
                        for block in globalData.bolckNameList:
                            flag2 = 0
                            for l in range(1, cols):
                                s = sheet.cell_value(0, l)
                                if (re.search(s, block)):
                                    ltemList.append(sheet.cell_value(globalData.HotSpots_N_List.index(item) + 1, l))
                                    flag2 = 1
                                    break
                            if (flag2 == 0):
                                ltemList.append(0)
                        dataListfortableWidget_5.append(ltemList)

                flagforwid_7 = 0

                # self.tableWidget_2.resizeColumnsToContents()
                for rowIndex in range(len(dataListfortableWidget_2)):
                    rowData = dataListfortableWidget_2[rowIndex]
                    if (flagforwid_7 == 0):
                        self.tableWidget_2.insertRow(rowIndex)
                        self.tableWidget_7.insertRow(flagforwid_7)
                        for j in range(len(rowData)):
                            item = QTableWidgetItem(str(dataListfortableWidget_2[rowIndex][j]))
                            item.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                            self.tableWidget_2.setItem(rowIndex, j, item)
                            item_7 = QTableWidgetItem(str(0))
                            item.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                            self.tableWidget_7.setItem(0, j, item_7)
                        flagforwid_7 = 1
                        self.tableWidget_7.cellChanged.connect(self.calculate)
                    else:
                        self.tableWidget_2.insertRow(rowIndex)
                        for j in range(len(rowData)):
                            item = QTableWidgetItem(str(dataListfortableWidget_2[rowIndex][j]))
                            item.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                            self.tableWidget_2.setItem(rowIndex, j, item)
                for rowIndex in range(len(dataListfortableWidget_4)):
                    rowData = dataListfortableWidget_4[rowIndex]
                    self.tableWidget_4.insertRow(0)
                    for j in range(len(rowData)):
                        item = QTableWidgetItem(str(dataListfortableWidget_4[rowIndex][j]))
                        item.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                        self.tableWidget_4.setItem(rowIndex, j, item)
                # print(dataListfortableWidget_5)
                for rowIndex in range(len(dataListfortableWidget_5)):
                    rowData = dataListfortableWidget_5[rowIndex]
                    self.tableWidget_5.insertRow(rowIndex)
                    self.tableWidget_6.insertRow(rowIndex)
                    for j in range(len(rowData)):
                        item = QTableWidgetItem(str(dataListfortableWidget_5[rowIndex][j]))
                        item.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                        self.tableWidget_5.setItem(rowIndex, j, item)
                    # self.tableWidget_6.setItem(rowIndex, 0, QTableWidgetItem(str('0K')))
                self.calculate(0, 0)

    def calculate(self, row, col):
        if (not self.tableWidget_7.item(row, col).text().isdigit()):
            self.messageDialog('警告', '只能输入数字')

            self.tableWidget_7.setItem(row, col, QTableWidgetItem(str(0)))

        else:
            ratioList = []
            tableWidget_6_ratioList = []
            tableWidget_7_List = []
            cyclesum = 0
            for i in range(globalData.BlockCount):
                if self.tableWidget_7.item(0, i):
                    cyclesum += int(self.tableWidget_7.item(0, i).text())
                    tableWidget_7_List.append(int(self.tableWidget_7.item(0, i).text()))
                else:
                    cyclesum += int(0)
                    tableWidget_7_List.append(int(0))

            for rowIndex in range(len(globalData.HotSpots_N_List)):
                sumRadio = 0
                for index in range(globalData.BlockCount):
                    y = int(tableWidget_7_List[index])
                    if y != 0:
                        if self.tableWidget_2.item(rowIndex, index):
                            a = self.tableWidget_2.item(rowIndex, index).text()
                            if (a != '0'):
                                li = a.split('e')
                                if (len(li) == 2):
                                    e = float(li[0])
                                    f = int(li[1])
                                    x = e * pow(10, f)
                                # x = f"{e * pow(10, f):.2e}"
                                else:
                                    x = float(li[0])
                            else:
                                x = float(0)
                            # print(tableWidget_7_List)
                            t = x * y
                            sumRadio += t
                ratioList.append(sumRadio)
            globalDataRatioList = []
            for i in range(len(ratioList)):
                globalDataRatioList.append(ratioList[i])
            lowRadio = 0
            if globalDataRatioList:
                globalDataRatioList.sort(reverse=True)
                if len(globalDataRatioList) >= 3:
                    lowRadio = globalDataRatioList[2]
                else:
                    lowRadio = globalDataRatioList[len(globalDataRatioList) - 1]
            rows = self.tableWidget_5.rowCount()
            for row in range(rows):
                sum = 0
                for index in range(globalData.BlockCount):
                    y = int(tableWidget_7_List[index])
                    if y != 0:
                        if self.tableWidget_5.item(row, index):
                            a = self.tableWidget_5.item(row, index).text()
                            if (a != '0'):
                                li = a.split('e')
                                if (len(li) == 2):
                                    e = float(li[0])
                                    f = int(li[1])
                                    x = e * pow(10, f)
                                else:
                                    x = float(li[0])
                            else:
                                x = float(0)
                            t = x * y
                            sum += t
                tableWidget_6_ratioList.append(sum)
            for i in range(len(tableWidget_6_ratioList)):
                if int(tableWidget_6_ratioList[i]) < lowRadio:
                    tableWidget_6_ratioList[i] = 'OK'
                else:
                    tableWidget_6_ratioList[i] = 'Nok'

            if (globalData.firstFlag == 0):
                for i in range(len(ratioList)):
                    self.tableWidget_3.insertRow(i)
                    item = QTableWidgetItem(str(ratioList[i]))
                    item.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                    self.tableWidget_3.setItem(i, 0, item)
                globalData.firstFlag = 1
            else:
                for i in range(len(ratioList)):
                    item = QTableWidgetItem(str(ratioList[i]))
                    item.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                    self.tableWidget_3.setItem(i, 0, item)
                for i in range(len(tableWidget_6_ratioList)):
                    item = QTableWidgetItem(str(tableWidget_6_ratioList[i]))
                    item.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                    self.tableWidget_6.setItem(i, 0, item)
            if (self.tableWidget_8.rowCount() == 0):
                self.tableWidget_8.insertRow(0)
            self.tableWidget_8.setItem(0, 0, QTableWidgetItem(str(cyclesum)))
            self.tableWidget_8.setVerticalHeaderLabels(['TOTAL = '])

    def HorSectionClicked(self, index):
        if len(globalData.HotSpots_N_List) != 0:
            blockName = globalData.bolckNameList[index]
            ans = getxlsxData(globalData.file2path, blockName)
            xlsxPath = ans[2]
            globalData.left = ans[0]
            globalData.right = ans[1]
            workbook = xlrd.open_workbook(xlsxPath)
            sheet = workbook.sheet_by_index(0)
            rows = sheet.nrows
            cols = sheet.ncols
            dialog = ChildWin_block()
            self.tableWidget_block = dialog.tableWidget
            dialog.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            dialog.tableWidget.setColumnCount(cols)
            dialog.tableWidget.horizontalHeader().setVisible(False)  # 隐藏水平表头
            print(sheet.row(1))
            titleList = sheet.row(1)

            if (str(sheet.cell_value(1, 4)) == str(sheet.cell_value(1, len(titleList) - 1))):
                globalData.GVWflag = len(titleList) - 1
            print(globalData.GVWflag)
            for i in range(rows):
                row = dialog.tableWidget.rowCount()
                dialog.tableWidget.insertRow(row)
                for j in range(cols):
                    item = QTableWidgetItem(str(sheet.cell_value(i, j)))
                    item.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                    dialog.tableWidget.setItem(row, j, item)
            dialog.tableWidget.setSpan(0, 4, 1, cols)
            dialog.tableWidget.cellChanged.connect(self.updateData)
            dialog.show()

            def pB_OK():
                dialog.close()
                os.remove(xlsxPath)

            # # 关联【确定】按钮

            dialog.buttonBox.clicked.connect(pB_OK)
            dialog.exec_()

        else:
            self.messageDialog('警告', '请先导入文件数据')

    def updateData(self, row, col):

        globalData.editChangeFlag = 1
        value = self.tableWidget_block.item(row, col).text()
        wb = load_workbook(filename=globalData.xlsxFile2path)
        ws = wb['Sheet1']

        if globalData.GVWflag != 0:
            if col == 4:
                if int(globalData.left) >= 26:
                    t = int(globalData.left) - 26
                    x = chr(t + col - 4 + (int(globalData.right) - int(globalData.left)) - 1 + 65)
                    x = 'A' + x
                    item = x + str(row + 1)
                    ws[item] = value
                else:
                    x = chr(int(globalData.left) + col - 4 + (int(globalData.right) - int(globalData.left)) - 1 + 65)
                    item = x + str(row + 1)
                    ws[item] = value
            elif col == globalData.GVWflag:
                if int(globalData.left) >= 26:
                    t = int(globalData.left) - 26
                    x = chr(t + col - 4 - (int(globalData.right) - int(globalData.left)) + 1 + 65)
                    x = 'A' + x
                    item = x + str(row + 1)
                    ws[item] = value
                else:
                    x = chr(int(globalData.left) + col - 4 - (int(globalData.right) - int(globalData.left)) + 1 + 65)
                    item = x + str(row + 1)
                    ws[item] = value

        if int(globalData.left) >= 26:
            t = int(globalData.left) - 26
            x = chr(t + col - 4 + 65)
            x = 'A' + x
            item = x + str(row + 1)
            ws[item] = value
        else:
            x = chr(int(globalData.left) + col - 4 + 65)
            item = x + str(row + 1)
            ws[item] = value
        if is_used(globalData.xlsxFile2path):
            self.messageDialog('警告', '文件正在被占用')
        else:
            wb.save(globalData.xlsxFile2path)
            xlsx_to_csv(globalData.xlsxFile2path)
            os.remove(globalData.xlsxFile2path)

    def noticeSave(self):
        if (globalData.editChangeFlag == 0):
            self.messageDialog('提示', '载荷数据没有做修改！')
        else:
            directory = QtWidgets.QFileDialog.getExistingDirectory(None, "选取文件夹", "C:/")  # 起始路径
            print(directory)

            shutil.copy(globalData.file2path, directory)
            self.messageDialog('提示', '保存成功！')

    def writetoCSV(self):
        file1 = globalData.file1path
        list = file1.rsplit("/", 1)
        filename = list[1]
        filepath = list[0]
        xlsxname = csv_to_xlsx_pd(globalData.cacheFilepath, filepath, filename)
        ansXlsxpath = os.path.join(globalData.cacheFilepath, 'cache', xlsxname).replace('\\', '/')
        wb = load_workbook(filename=ansXlsxpath)
        fill = PatternFill("solid", fgColor="8064a2")
        fill2 = PatternFill("solid", fgColor="ffeb9c")
        ws = wb['Sheet1']
        for i in range(len(globalData.bolckNameList)):
            x = chr(i + 68)
            item = x + str(1)
            ws[item] = globalData.bolckNameList[i]
            ws[item].fill = fill
        ws[chr(len(globalData.bolckNameList) + 68) + str(1)] = 'RADIO'
        ws[chr(len(globalData.bolckNameList) + 68) + str(1)].fill = fill2
        rows_2 = self.tableWidget_2.rowCount()
        cols_2 = self.tableWidget_2.columnCount()
        for rowindex in range(rows_2):
            for colindex in range(cols_2):
                item = chr(colindex + 68) + str(rowindex + 2)
                ws[item] = self.tableWidget_2.item(rowindex, colindex).text()
                if (colindex == cols_2 - 1):
                    item_3 = chr(cols_2 + 68) + str(rowindex + 2)
                    ws[item_3] = self.tableWidget_3.item(rowindex, 0).text()
                    ws[item_3].fill = fill2
        interval = 4

        # tableWidget_4
        rows_4 = self.tableWidget_4.rowCount()
        cols_4 = self.tableWidget_4.columnCount()
        for rowindex in range(rows_4):
            for colindex in range(cols_4):
                item = chr(colindex + 65) + str(rowindex + interval + rows_2)
                if self.tableWidget_4.item(rowindex, colindex):
                    x = self.tableWidget_4.item(rowindex, colindex).text()
                    ws[item] = x
                # if (colindex == cols_4 - 1):
                #     item_3 = chr(cols_4 + 65) + str(rowindex + interval + rows_2)
                #     ws[item_3] = self.tableWidget_6.item(rowindex, 0).text()
                #     ws[item_3].fill = fill2

        # tableWidget_5
        rows_5 = self.tableWidget_5.rowCount()
        cols_5 = self.tableWidget_5.columnCount()
        for rowindex in range(rows_5):
            for colindex in range(cols_5):
                item = chr(colindex + 68) + str(rowindex + interval + rows_2)
                ws[item] = self.tableWidget_5.item(rowindex, colindex).text()
                if (colindex == cols_5 - 1):
                    item_3 = chr(cols_5 + 68) + str(rowindex + interval + rows_2)
                    ws[item_3] = self.tableWidget_6.item(rowindex, 0).text()
                    ws[item_3].fill = fill2
        # tableWidget_7
        cols_7 = self.tableWidget_7.columnCount()

        for colindex in range(cols_7):
            item = chr(colindex + 68) + str(interval * 2 + rows_2 + rows_5)
            ws[item] = self.tableWidget_7.item(0, colindex).text()
            if (colindex == cols_7 - 1):
                item_3 = chr(cols_7 + 68) + str(interval * 2 + rows_2 + rows_5)
                ws[item_3] = self.tableWidget_8.item(0, 0).text()
                ws[item_3].fill = fill2
        wb.save(ansXlsxpath)
        directory = QtWidgets.QFileDialog.getExistingDirectory(None, "选取文件夹", "C:/")  # 起始路径
        if directory:
            datetime_object = datetime.datetime.now()
            time = str(datetime_object).split(' ')[0]
            list = ansXlsxpath.rsplit("/", 1)
            filename = 'Result_IterationN_' + time + '.xlsx'
            filepath = list[0]
            saveFile = os.path.join(filepath, filename).replace('\\', '/')

            if os.path.exists(saveFile):
                os.remove(saveFile)
            # elif is_used(saveFile):
            #     self.messageDialog('警告', '文件正在被占用')
            else:
                os.rename(ansXlsxpath, saveFile)
                shutil.copy(saveFile, directory)
                self.messageDialog('提示', '保存成功！')

    def messageDialog(self, type, message):
        # 核心功能代码就两行，可以加到需要的地方
        if type == '警告':
            x = QMessageBox.Warning
        if type == '提示':
            x = QMessageBox.Information
        msg_box = QMessageBox(x, type, message)
        msg_box.exec_()


# if __name__ == "__main__":
def main():
    app = QApplication(sys.argv)
    # setup stylesheet
    apply_stylesheet(app, theme='light_blue.xml', invert_secondary=True)
    myWin = MyMainForm()
    myWin.show()
    sys.exit(app.exec_())
