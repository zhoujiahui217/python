# -*- coding: utf-8 -*-

import xlrd
from xlutils.copy import copy
import sys

sys.path.append("..")
import logging.config


class ExcelManage:
    def __init__(self):
        pass

    def read_excel(self, url, w):
        """读取excel"""
        self.w = w
        book = xlrd.open_workbook(url)
        self.sheet = book.sheet_by_name("Sheet1")

        oldWb = xlrd.open_workbook(url)
        self.w = copy(oldWb)
        return self.w

    def write_excel(self, result, reason, file):
        """向excel中写入结果"""
        logging.info('write excel')
        try:
            self.w.get_sheet(0).write(self.len, 5, result)
            self.w.get_sheet(0).write(self.len, 6, reason)
            self.w.save(file)
            print 'success'

        except IOError:
            print 'error'
            logging.exception('Exception Logged')

ex=ExcelManage()
ex.wr