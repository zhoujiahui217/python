#! /usr/bin/python
# -*- coding: utf-8 -*-
from selenium import webdriver
import time
import xlrd
import datetime
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
import selenium.webdriver.support.expected_conditions as EC
import selenium.webdriver.support.ui as ui
from xlwt import *
from xlutils.copy import copy
from selenium.webdriver.common.keys import Keys
import json
import urllib
import urllib2


class compare():
    def openfile(self, url, times, file):
        """wait elements"""

        def is_visible(locator, timeout=10):
            try:
                ui.WebDriverWait(self.browser, timeout).until(EC.visibility_of_element_located((By.XPATH, locator)))
                False
                return True
            except TimeoutException:
                return

        book = xlrd.open_workbook(url)
        sheet = book.sheet_by_name("Sheet1")

        oldWb = xlrd.open_workbook(url)
        self.w = copy(oldWb)
        self.browser = webdriver.Firefox()
        self.browser.get("file:///C:/Users/lenovo/Desktop/8.html")
        time.sleep(2)
        self.browser.find_element_by_xpath('html/body/div[2]/a').click()
        frame = self.browser.find_element_by_xpath("html/body/div/div/iframe")
        time.sleep(2)
        self.browser.switch_to.frame(frame)

        for i in range(0, times):
            print i

            time.sleep(2)

            self.num = i  # 插入excel行数
            row_data = sheet.row_values(i)
            print row_data[0]

            self.realanswer = row_data[2]  # excel中问题的准确答案

            self.realqustion = row_data[1]  # excel中问题的相似答案
            time.sleep(2)
            self.browser.find_element_by_id("msg").send_keys(row_data[0])

            is_visible(".//*[@id='send']")
            self.browser.find_element_by_xpath(".//*[@id='send']").click()
            time.sleep(2)
            self.starttime = datetime.datetime.now()  # 开始时间

            try:

                self.leng = self.browser.find_element_by_xpath(".//*[@id='msgList']/li[last()]/div[2]/div").text.split(
                    '\n')

                self.endtime = datetime.datetime.now()  # 结束时间
                self.robotanswer = self.browser.find_element_by_xpath(
                    ".//*[@id='msgList']/li[last()]/div[2]/div/div/div/div/ol/li[1]/a")

                usetime = self.endtime - self.starttime

                # self.leng = self.browser.find_element_by_xpath(
                #    ".//*[@id='msgList']/li[last()]/div[2]/div/").text.split('\n')
                # print self.text
                # self.leng = self.browser.find_element_by_xpath(".//*[@class='clearfix magBox leftMsg'][last()]/div/div/div/div/div/ul").text.split('\n')

                length = len(self.leng)
                tag = '、'.decode('utf-8')
                for n in range(1, int(length)):
                    #is_visible(".//*[@id='msgList']/li[last()]/div[2]/div/div/div/div/ol/li[" + str(n) + "]/a")
                    self.robotanswer = self.browser.find_element_by_xpath(
                        ".//*[@id='msgList']/li[last()]/div[2]/div/div/div/div/ol/li[" + str(n) + "]/a").text.split(
                        tag)[1]
                    #time.sleep(0.5)
                    self.robotanswer = self.robotanswer.replace(" ", "")
                    self.realqustion =self.realqustion.replace(" ", "")
                    if self.robotanswer == self.realqustion:
                        res = 'ok'
                        print res
                        break
                    else:
                        res = 'error'
                        print res

                self.writeexcel(" ", " ", " ", file)

            except Exception, e:
                # print 'str(Exception):\t', str(Exception)
                # print 'str(e):\t\t', str(e)
                try:

                    self.robotanswer = self.browser.find_element_by_xpath(
                        ".//*[@id='msgList']/li[@class='robot-say'][last()]/div[2]/div/div/div").text  # 机器回答的答案

                    self.endtime = datetime.datetime.now()  # 结束时间

                    usetime = self.endtime - self.starttime
                    # print self.answer
                    self.realanswer = self.realanswer.replace("\n", "").replace(" ", "")
                    self.robotanswer = self.robotanswer.replace("\n", "").replace(" ", "")

                    print self.realanswer
                    print self.robotanswer

                    if self.robotanswer == self.realanswer:
                        res = 'ok'
                        print res
                    else:
                        res = 'error'
                        print res

                    self.writeexcel(str(usetime), 1, res, file)

                    # self.browser.quit()
                except Exception, e:
                    # print 'str(Exception):\t', str(Exception)
                    # print 'str(e):\t\t', str(e)
                    self.writeexcel(0, 0, 'error', file)

                    # self.browser.quit()

    def writeexcel(self, ut, length, result, file):
        """向excel中写入结果"""

        self.w.get_sheet(0).write(self.num, 3, ut)
        self.w.get_sheet(0).write(self.num, 4, length)
        self.w.get_sheet(0).write(self.num, 5, result)
        self.w.save(file)
        print 'success'


if __name__ == '__main__':
    com = compare()
    # com.openfile(r'D:\compare\30\3.xlsx', 269, r'D:\compare\30\zc1_res.xls')
    # com.openfile(r'C:\Users\lenovo\Documents\old\3.29-1.xlsx', 200, r'D:\compare\aljust\aljust_291.xls')
    com.openfile(r'C:\Users\lenovo\Documents\old\3.1.xlsx', 250, r'D:\compare\aljust2\aijust_28.2.xls')
