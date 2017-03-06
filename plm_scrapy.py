#!/usr/bin/env python
# -*- coding:cp936 -*-
import time
import os
import codecs
import selenium.common.exceptions
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import tkMessageBox


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
filename = codecs.open('个人PLM上传文档数量.csv'.decode('gbk'), 'w', 'gbk')


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

for index, item in enumerate(file_doc_list):
    filename.write(item + "," + file_time_list[index] + os.linesep)
length_filedoc_list = len(file_doc_list)
print length_filedoc_list
filename.close()
browser.quit()
print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
tkMessageBox.showinfo('完成啦'.decode('gbk'), '数据抓取完毕，共抓取到%s个结果！保存在《个人PLM上传文档数量.csv》中'.decode('gbk') %length_filedoc_list)
