# -*- coding: utf-8 -*-

# import logging
import json
# import requests
# import common.excel
from common.rate import *
import xlrd
import datetime
from xlutils.copy import copy
import sys
# import common.use_mysql
import common.common_fun

sys.path.append("..")
import logging.config

reload(sys)
sys.setdefaultencoding('utf-8')


class GetAnswer:
    def __init__(self, read_url, save_url, start_time='', end_time="", use_time=''):

        self.READ_URL = read_url  # 读取文件路径
        self.SAVE_URL = save_url  # 保存文件路径

        self.start_time = start_time
        self.end_time = end_time
        self.use_time = use_time

    def read_excel(self, url, sheet_name):

        """读取excel"""
        self.sheet_name = sheet_name
        book = xlrd.open_workbook(url)
        self.sheet = book.sheet_by_name(sheet_name)
        oldWb = xlrd.open_workbook(self.READ_URL)
        self.w = copy(oldWb)

    def write_excel(self, time, top, result, reason, file):

        """向excel中写入结果"""

        try:

            if self.sheet_name != 'fanli':
                self.w.get_sheet(0).write(self.len, 6, time)
                self.w.get_sheet(0).write(self.len, 8, top)
                self.w.get_sheet(0).write(self.len, 9, result)
                self.w.get_sheet(0).write(self.len, 10, reason)
            else:
                self.w.get_sheet(1).write(self.len, 6, time)
                self.w.get_sheet(1).write(self.len, 8, top)
                self.w.get_sheet(1).write(self.len, 9, result)
                self.w.get_sheet(1).write(self.len, 10, reason)

            self.w.save(file)
            # print 'success'

        except IOError:
            print 'error'

    def send_msg(self, times=0, func='', sh='Sheet1'):

        # self.func = func

        self.read_excel(self.READ_URL, sh)
        for n in range(0, times):
            self.len = n
            row_data = self.sheet.row_values(n)
            self.main_question = row_data[1]
            self.real_answer = row_data[2]
            self.like_question = row_data[0]
            logging.info(self.like_question)
            # print self.like_question

            self.start_time = datetime.datetime.now()  # 开始时间

            self.rest = common.common_fun.send_msg(self.like_question)
            print self.rest
            self.end_time = datetime.datetime.now()  # 结束时间
            self.use_time = self.end_time - self.start_time

            func()

    def do_score(self):
        try:

            value = self.rest['debug-info']['final_result']['value']
            #
            #
            #
            if len(value) == 0:
                self.write_excel(str(self.use_time), len(value), "Fail", "no answer", self.SAVE_URL)
                self.w.get_sheet(0).write(self.len, 7, '0')
                print 'no'
            for i in range(len(value)):
                if len(value) == 1:
                    score = value[i]
                    # extend_id = score[2]['value']
                    correct_question = score[3]['value']
                    # sql语句
                    # mysql_str = "select question FROM main_question WHERE tenant_id=46 AND  is_delete=0 AND id in (select main_id FROM like_question WHERE question='" + correct_question + "' and tenant_id=46);"
                    # """调用mysql"""
                    # info = common.use_mysql.get_mysql('222.180.162.163','reader','test001','intelligent_server',mysql_str)
                    # try:
                    #
                    #     que = info[0][0]
                    #
                    # except:
                    #     que = ""
                    ans = score[8]['value']
                    answer_json=json.loads(ans)
                    answer=answer_json['answer']
                    self.real_answer = self.real_answer.replace("\n", "").replace(" ", "")
                    answer = answer.replace("\n", "").replace(" ", "")
                    # a = json.loads(answer)
                    # a = a['answer']
                    # a = a.replace("\n", "").replace(" ", "")
                    print 'right_answer' + answer
                    if answer == self.real_answer or correct_question == self.main_question:

                        # and correct_question == self.main_question
                        self.write_excel(str(self.use_time), 1, "pass", 1, self.SAVE_URL)
                        self.w.get_sheet(0).write(self.len, 3, correct_question)

                        self.w.get_sheet(0).write(self.len, 7, '1')
                        # self.w.get_sheet(0).write(self.len, 12, que)
                    else:

                        self.write_excel(str(self.use_time), 1, "Fail", " answer false ", self.SAVE_URL)
                        self.w.get_sheet(0).write(self.len, 3, correct_question)
                        self.w.get_sheet(0).write(self.len, 7, '1')
                        # self.w.get_sheet(0).write(self.len, 12, que)

                elif len(value) != 1 and len(value) != 0:

                    score = value[i]
                    correct_question = score[3]['value']
                    answer = score[8]['value']
                    a=json.loads(answer)
                    #a=a['answer']
                    answer = answer.replace("\n", "").replace(" ", "")
                    """临时修改返回参数为json数据的格式"""
                    #a = json.loads(answer)
                    a = a['answer'].replace("\n", "").replace(" ", "")
                    # print a['answer']
                    self.real_answer=self.real_answer.replace("\n", "").replace(" ", "")
                    if correct_question == self.main_question or self.real_answer == a:  # if correct_question == self.main_question or self.real_answer == answer:

                        self.write_excel(str(self.use_time), 1, "pass", int(i) + 1, self.SAVE_URL)
                        self.w.get_sheet(0).write(self.len, 3 + i, correct_question)
                        self.w.get_sheet(0).write(self.len, 7, 'n')
                        break

                    else:
                        self.write_excel(str(self.use_time), 0, "Fail", int(i) + 1, self.SAVE_URL)
                        # logging.debug('top3 error')
                        self.w.get_sheet(0).write(self.len, 3 + i, correct_question)
                        self.w.get_sheet(0).write(self.len, 7, 'n')
                        # break

                else:
                    self.write_excel(str(self.use_time), len(value), "Fail", "无答案 ", self.SAVE_URL)
                    self.w.get_sheet(0).write(self.len, 3 + i, correct_question)
                    self.w.get_sheet(0).write(self.len, 7, '')
                    print 'no'
        except:
            self.write_excel(str(self.use_time), 0, "Fail", "unknown", self.SAVE_URL)

    def base_info(self):
        value = self.rest['debug-info']['base_info']['value']
        if len(value) != 0:
            info = value[0]
            print info
            print info[1]
            if info[1]['value'] == 'meta-name_rules':
                self.write_excel(str(self.use_time), len(value), "pass", "", self.SAVE_URL)
            else:
                self.write_excel(str(self.use_time), len(value), "pass", "", self.SAVE_URL)
        else:
            self.write_excel(str(self.use_time), len(value), "error", "", self.SAVE_URL)

        self.i = 0

        self.read_excel(self.SAVE_URL, 'Sheet1')
        num = self.sheet.nrows
        print num

    def es_info(self):
        try:
            value = self.rest['debug-info']['es']['value']

            if len(value) == 0:
                self.write_excel(str(self.use_time), 0, "error", "no answer", self.SAVE_URL)
            for i in range(len(value)):

                str_info_out = value[i][8]['value']
                self.main_question = self.main_question.replace('\n', '').replace(" ", "")
                str_info_out = str_info_out.replace('\n', '').replace(" ", "")

                if self.main_question == str_info_out or self.like_question == str_info_out:
                    if i == 0:
                        self.write_excel(str(self.use_time), 1, "pass", i + 1, self.SAVE_URL)
                    elif i <= 2:
                        self.write_excel(str(self.use_time), 1, "pass", i + 1, self.SAVE_URL)
                    else:
                        self.write_excel(str(self.use_time), 1, "pass", "", self.SAVE_URL)

                    # self.write_excel(str(self.use_time), 1, "pass", "", self.SAVE_URL)
                    break
                else:
                    self.write_excel(str(self.use_time), 0, "error", "", self.SAVE_URL)
        except:
            self.write_excel(str(self.use_time), 0, "error", "no answer", self.SAVE_URL)

    def wmd_info(self):
        try:
            value = self.rest['debug-info']['topN']['value']

            ss = datetime.datetime.now()
            # print ss - self.start_time
            # print len(value)
            if len(value) == 0:
                self.write_excel(str(self.use_time), 0, "error", "no answer", self.SAVE_URL)
            else:
                for w in range(0, 3):  #
                    que = value[w][3]['value']
                    asw = value[w][8]['value']

                    asw_json = json.loads(asw)
                    answer = asw_json['answer']
                    self.real_answer = self.real_answer.replace('\n', '').replace(" ", "")
                    answer = answer.replace("\n", "").replace(" ", "")
                    if self.main_question == que or self.real_answer == answer:
                        if w == 0:
                            self.write_excel(str(self.use_time), 1, "pass", w + 1, self.SAVE_URL)
                            # self.sheet_name
                        elif w <= 2:
                            self.write_excel(str(self.use_time), 1, "pass", w + 1, self.SAVE_URL)

                        else:
                            self.write_excel(str(self.use_time), 0, "error", "over 3", self.SAVE_URL)

                        break
                    else:
                        self.write_excel(str(self.use_time), 0, "error", "", self.SAVE_URL)
            e = datetime.datetime.now()
        except:
            self.write_excel(str(self.use_time), 0, "no answer", "", self.SAVE_URL)
            # e - self.start_time


if __name__ == "__main__":
    logging.config.fileConfig(r'../config/logging.config')

    ga = GetAnswer(r'../data/11.xlsx', r'../report/test_do_1017.xls')

    ga.send_msg(1181, ga.do_score, 'zhengli')
    # ga.send_msg(10, ga.base_info)
    # ga.send_msg(309, ga.es_info, 'zhengli')
    #ga.send_msg(309, ga.wmd_info, 'zhengli')
    r = Rate()
    r.rate(r'../report/test_do_1017.xls', 'ga.do_score', 'zhengli')
