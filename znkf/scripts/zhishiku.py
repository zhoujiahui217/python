# -*- coding: utf-8 -*-

import xlrd
from xlutils.copy import copy
import sys
from pyExcelerator import *

sys.path.append("..")


class zhishiku:
    def __init__(self):
        self.main_question_list = []
        self.like_question_list = []
        self.answer_list = []
        self.mm = []

        self.read_excel('../data/all0919.xls')
        for x in range(0, 2081):
            row = self.sheet.row_values(x)
            self.mm.append(row[0])

    def read_excel(self, url):
        """读取excel"""
        # self.w = w
        book = xlrd.open_workbook(url)
        self.sheet = book.sheet_by_name("Sheet1")

    def write_excel(self):
        """向excel中写入结果"""
        self.read_excel('../data/all0919.xls')
        len = self.sheet.nrows
        book = xlrd.open_workbook(r'../data/all0919.xls')
        self.sheet_2 = book.sheet_by_name("Sheet1")

        oldWb = xlrd.open_workbook(r'../data/all0919.xls')
        self.w = copy(oldWb)

    def del_repeat(self):
        """去重"""
        self.write_excel()
        for m in range(0, 2081):
            print m
            row = self.sheet.row_values(m)

            question = row[0]
            like_question = row[1]
            answer = row[2]
            # print self.answer_list
            # print self.like_question_list
            # print self.main_question_list
            if m == 0:
                self.answer_list.append('answer')
                self.like_question_list.append('like_question')
                self.main_question_list.append('main_question')

            if answer in self.answer_list and like_question in self.like_question_list and question in self.main_question_list:
                self.w.get_sheet(0).write(m, 0, '')
                self.w.get_sheet(0).write(m, 1, '')
                self.w.get_sheet(0).write(m, 2, '')

            else:
                if question != like_question:
                    question = question.replace('\n', '').replace(" ", "")

                    self.answer_list.append(answer)
                    self.like_question_list.append(like_question)
                    self.main_question_list.append(question)

                    self.w.get_sheet(0).write(m, 0, question)
                    self.w.get_sheet(0).write(m, 1, like_question)
                    self.w.get_sheet(0).write(m, 2, answer)
                    self.w.get_sheet(0).write(m, 4, 'ok')
                else:
                    self.w.get_sheet(0).write(m, 0, '')
                    self.w.get_sheet(0).write(m, 1, '')
                    self.w.get_sheet(0).write(m, 2, '')

            self.w.save(r'../report/19.xls')

    def del_likequestion(self):
        self.write_excel()

        for m in range(0, 2081):
            row2 = self.sheet.row_values(m)
            like_question = row2[1]
            if like_question in self.mm:
                self.w.get_sheet(0).write(m, 6, 'repeat')

        self.w.save(r'../report/19.xls')


    def change_excel(self):
        # self.write_excel()
        for i in range(0, 10):
            row1 = self.sheet.row_values(i)
            question1 = row1[0]
            l_question = row1[1]
            answer = row1[2]
            if i == 0:
                self.w.get_sheet(0).write(i, 0, question1)
                self.w.get_sheet(0).write(i, 1, l_question)
                self.w.get_sheet(0).write(i, 2, answer)
                self.w.save(r'../report/11.xls')
            else:
                for n in range(1, 10):
                    row2 = self.sheet.row_values(n)
                    question2 = row2[0]
                    l_question2 = row2[1]
                    answer2 = row2[2]
                    if row1 == row2:
                        self.w.get_sheet(0).write(i, 2 + n, l_question2)
                        self.w.get_sheet(0).write(n, 2 + n, '')
                    else:
                        self.w.get_sheet(0).write(i, 0, question2)
                        self.w.get_sheet(0).write(i, 1, l_question2)
                        self.w.get_sheet(0).write(i, 2, answer2)
                    i = i + n

                    self.w.save(r'../report/11.xls')


if __name__ == "__main__":
    c = zhishiku()
    c.del_likequestion()
