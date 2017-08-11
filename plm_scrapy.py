#!/usr/bin/env python
# -*- coding:cp936 -*-
import time
import os
import xlsxwriter
import selenium.common.exceptions
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import tkMessageBox
import datetime


print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
file_doc_list = []
file_time_list = []
firefoxdriverPath = os.path.abspath(os.path.curdir)
browser = webdriver.Firefox(firefoxdriverPath)
url = "http://plm.inspur.com/Windchill/app/"
browser.get(url)
WebDriverWait(browser, 3600).until(ec.presence_of_all_elements_located((By.XPATH, "//div[@class='x-grid3-cell-inner x-grid3-col-ASSIGNMENT_SUBJECT']")))
time.sleep(5)
scrollbar_test = browser.find_element_by_xpath("//div[contains(@id, 'wrap_user-homepage_tab_jcaGen')]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[2]")


def get_data():
    tables = browser.find_elements_by_css_selector("div.x-grid3-cell-inner.x-grid3-col-ASSIGNMENT_SUBJECT > a")
    browser.implicitly_wait(1)
    for abc in tables:
        filename_file = abc.get_attribute('ext:qtip').split(",")[1]
        submit_time = abc.find_element_by_xpath('parent::div/parent::td/parent::tr/td[13]/div')
        filename_role = abc.find_element_by_xpath('parent::div/parent::td/parent::tr/td[19]/div')
        filename_status = abc.find_element_by_xpath('parent::div/parent::td/parent::tr/td[11]/div')
        browser.implicitly_wait(1)
        filename_time = submit_time.get_attribute("ext:qtip").split(" ")[0]
        role_you = filename_role.get_attribute('ext:qtip')
        status_file = filename_status.get_attribute('ext:qtip')
        if filename_file not in file_doc_list and role_you == '提交者'.decode('gbk') and status_file == '已完成'.decode('gbk'):
            file_doc_list.append(filename_file)
            file_time_list.append(filename_time)


for i in range(0, 100):
    try:
        get_data()
    except selenium.common.exceptions.StaleElementReferenceException:
        get_data()
    scrollbar_test.send_keys(Keys.PAGE_DOWN)
    time.sleep(5)


timestamp = time.strftime('%Y%m%d', time.localtime())
workbook_to_write = xlsxwriter.Workbook("个人PLM上传文档数量-%s.xlsx".decode('gbk') % timestamp)
formatone = workbook_to_write.add_format()
formatone.set_border(1)
sheet_now = workbook_to_write.add_worksheet("个人PLM上传文档数量".decode('gbk'))
sheet_now.set_column('A:A', 100)
sheet_now.set_column('B:B', 25)
formattitle = workbook_to_write.add_format()
formattitle.set_border(1)
formattitle.set_align('center')
formattitle.set_bg_color("yellow")
formattitle.set_bold(True)

list_title = ["报告标题".decode('gbk'), "报告上传时间".decode('gbk')]

sheet_now.merge_range(0, 0, 0, 2, "个人PLM上传文档数量".decode('gbk'), formattitle)

for index_title, item_title in enumerate(list_title):
    sheet_now.write(1, index_title, item_title, formatone)

for index_data, item_data in enumerate(file_doc_list):
    sheet_now.write(index_data + 2, 0, item_data, formatone)
    sheet_now.write(index_data + 2, 1, datetime.datetime.strptime(file_time_list[index_data], '%Y/%m/%d'), workbook_to_write.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))
workbook_to_write.close()
length_filedoc_list = len(file_doc_list)
browser.quit()
print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
tkMessageBox.showinfo('完成啦'.decode('gbk'), '数据抓取完毕，共抓取到%s个结果！保存在《个人PLM上传文档数量-%s.xlsx》中'.decode('gbk') % (timestamp, length_filedoc_list))
