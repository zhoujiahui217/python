# -*- coding: utf-8 -*-

import xlrd
from xlutils.copy import copy
from pyExcelerator import *

sys.path.append("..")


class change:
    """format of excel  , test template change to zhishiku template  """

    def read_excel(self):
        """读取excel"""
        # self.w = w
        book = xlrd.open_workbook('../data/0920.xls')
        self.sheet = book.sheet_by_name("Sheet1")

        # oldWb = xlrd.open_workbook(url)
        # self.w = copy(oldWb)
        # return self.w

    def write_excel(self):
        """向excel中写入结果"""
        self.read_excel()
        self.length = self.sheet.nrows
        book = xlrd.open_workbook('../data/0920.xls')
        self.sheet_2 = book.sheet_by_name("Sheet1")

        oldWb = xlrd.open_workbook('../data/0920.xls')
        self.w = copy(oldWb)

    def change_excel(self):
        # self.write_excel()
        i = 0
        while i < self.length:
            print i
           
            row1 = self.sheet.row_values(i)
            question1 = row1[0]  # first line question's value
            l_question = row1[1]
            answer = row1[2]
            x = 0  # like question's num
            for n in range(i + 1, self.length):
                row2 = self.sheet.row_values(n)
                question2 = row2[0]  # next line question's value
                l_question2 = row2[1]
                answer2 = row2[2]

                if question1 == question2:

                    self.w.get_sheet(0).write(i, 3 + x, l_question2)

                    self.w.get_sheet(0).write(n, 0, '')
                    self.w.get_sheet(0).write(n, 1, '')
                    self.w.get_sheet(0).write(n, 2, '')
                    self.w.save(r'../report/20.xls')
                    x += 1

                else:
                    i = i + 1
                    break
            i = i + x
            self.w.save(r'../report/20.xls')


if __name__ == "__main__":
    c = change()
    c.write_excel()
    c.change_excel()
