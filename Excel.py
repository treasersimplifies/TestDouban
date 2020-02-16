import xlrd
import xlwt
from xlutils.copy import copy
import openpyxl
# from xlwt import Workbook

class Excel(object):

    def __init__(self, filePath, fileSavePath):
        self.filePath = filePath
        self.fileSavePath = fileSavePath
        self.wb = xlrd.open_workbook(filePath)
        self.sheet = self.wb.sheet_by_index(0)
        self.new_excel = copy(self.wb)
        idx = self.wb.sheet_names().index('测试用例')
        self.new_excel.get_sheet(idx).name = u'测试结果'

    def readColN(self, n):
        col = self.sheet.col_values(n)
        return col

    def readCases(self):
        # wb = xlrd.open_workbook(filePath)
        # testcases = wb.sheet_by_index(0)
        col = self.sheet.col_values(3)
        return col

    def writeResults(self, n1, n2, res):
        # old_excel = xlrd.open_workbook(filePath)
        ws = self.new_excel.get_sheet(0)
        ws.write(n1, n2, res)
        # self.new_excel.save(self.fileSavePath)
        return self.new_excel

    def saveResult(self):
        self.new_excel.save(self.fileSavePath)
    