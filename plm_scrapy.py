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
from selenium.webdriver.common.action_chains import ActionChains
import tkMessageBox
import datetime

print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
list_sn = []
list_filename_doc = []
list_version = []
list_date_created = []
list_link = []
list_location = []
list_feichu = []
list_type = []
firefoxdriverPath = os.path.abspath(os.path.curdir)
browser = webdriver.Firefox(firefoxdriverPath)
# driverpath = os.path.join(os.path.abspath(os.path.curdir), "chromedriver.exe")
# browser = webdriver.Chrome(driverpath)
# browser.maximize_window()
url = "http://plm.inspur.com/Windchill/app/"
browser.get(url)
WebDriverWait(browser, 3600).until(ec.presence_of_all_elements_located(
    (By.XPATH, "//div[@class='x-grid3-cell-inner x-grid3-col-ASSIGNMENT_SUBJECT']")))
time.sleep(5)
scrollbar_test = browser.find_element_by_xpath(
    "//div[contains(@id, 'wrap_user-homepage_tab_jcaGen')]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[2]")


def get_data():
    tables = browser.find_elements_by_css_selector("div.x-grid3-cell-inner.x-grid3-col-ASSIGNMENT_SUBJECT > a")
    browser.implicitly_wait(1)
    for abc in tables:
        sn = abc.get_attribute('ext:qtip').split(",")[0].split("-")[1].strip()
        filename_file = abc.get_attribute('ext:qtip').split(",")[1].strip()
        version = abc.get_attribute('ext:qtip').split(",")[-1].strip()
        link_all = abc.get_attribute('href').strip()
        submit_time = abc.find_element_by_xpath('parent::div/parent::td/parent::tr/td[13]/div')
        filename_role = abc.find_element_by_xpath('parent::div/parent::td/parent::tr/td[19]/div')
        filename_status = abc.find_element_by_xpath('parent::div/parent::td/parent::tr/td[11]/div')
        browser.implicitly_wait(1)
        filename_time = submit_time.get_attribute("ext:qtip").split(" ")[0].strip()
        role_you = filename_role.get_attribute('ext:qtip').strip()
        status_file = filename_status.get_attribute('ext:qtip').strip()
        if sn not in list_sn and role_you == '提交者'.decode('gbk') and status_file == '已完成'.decode('gbk'):
            list_sn.append(sn)
            list_filename_doc.append(filename_file)
            list_version.append(version)
            list_date_created.append(filename_time)
            list_link.append(link_all)


for i in range(0, 50):
    print i
    try:
        get_data()
    except selenium.common.exceptions.StaleElementReferenceException:
        get_data()
    scrollbar_test.send_keys(Keys.PAGE_DOWN)
    time.sleep(5)

for index_link, item_link in enumerate(list_link):
    newwindow = 'window.open("%s");' % item_link
    browser.execute_script(newwindow)
    time.sleep(2)
    handles = browser.window_handles
    browser.switch_to.window(handles[-1])
    WebDriverWait(browser, 30).until(ec.presence_of_all_elements_located((By.CSS_SELECTOR, "div#dataStoreMoreAttributesGroup > table > tbody > tr:nth-child(3) > td:nth-child(6) > a")))
    list_location_temp = browser.find_elements_by_css_selector(
        "div#dataStoreMoreAttributesGroup > table > tbody > tr:nth-child(3) > td:nth-child(6) > a")
    list_location_src_temp = []
    for item_temp in list_location_temp:
        data_location_temp = item_temp.text.strip()
        list_location_src_temp.append(data_location_temp)
    data_location_to_write = "-".join(list_location_src_temp)
    list_location.append(data_location_to_write)

    try:
        data_feichu = browser.find_element_by_css_selector("div#dataStoreBusinessAttributesGroup > table > tbody > tr:nth-child(1) > td:nth-child(3)").text.strip()
        if len(data_feichu) == 0:
            list_feichu.append("None")
        else:
            list_feichu.append(data_feichu)

        data_type = browser.find_element_by_css_selector(
            "div#dataStoreBusinessAttributesGroup > table > tbody > tr:nth-child(2) > td:nth-child(3)").text.strip()
        list_type.append(data_type)
    except selenium.common.exceptions.NoSuchElementException:
        list_feichu.append("None")
        list_type.append("None")

    browser.close()
    browser.switch_to.window(handles[0])

browser.quit()

timestamp = time.strftime('%Y%m%d', time.localtime())
workbook_to_write = xlsxwriter.Workbook("个人PLM上传文档数量-%s.xlsx".decode('gbk') % timestamp)
formatone = workbook_to_write.add_format()
formatone.set_border(1)
sheet_now = workbook_to_write.add_worksheet("个人PLM上传文档数量".decode('gbk'))
sheet_now.set_column('A:A', 15)
sheet_now.set_column('B:B', 55)
sheet_now.set_column('C:C', 10)
sheet_now.set_column('D:E', 15)
sheet_now.set_column('F:F', 35)
sheet_now.set_column('G:G', 25)

formattitle = workbook_to_write.add_format()
formattitle.set_border(1)
formattitle.set_align('center')
formattitle.set_bg_color("yellow")
formattitle.set_bold(True)

list_title = ["报告编号".decode('gbk'), "报告标题".decode('gbk'), "报告版本".decode('gbk'), "报告上传时间".decode('gbk'),
              "报告分类".decode('gbk'), "报告所在位置".decode('gbk'), "替换废除文件编号".decode('gbk'), ]

sheet_now.merge_range(0, 0, 0, 6, "个人PLM上传文档数量".decode('gbk'), formattitle)

for index_title, item_title in enumerate(list_title):
    sheet_now.write(1, index_title, item_title, formatone)

for index_data, item_data in enumerate(list_sn):
    sheet_now.write(index_data + 2, 0, item_data, formatone)
    sheet_now.write(index_data + 2, 1, list_filename_doc[index_data], formatone)
    sheet_now.write(index_data + 2, 2, list_version[index_data], formatone)
    sheet_now.write(index_data + 2, 3, datetime.datetime.strptime(list_date_created[index_data], '%Y/%m/%d'),
                    workbook_to_write.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))
    sheet_now.write(index_data + 2, 4, list_type[index_data], formatone)
    sheet_now.write(index_data + 2, 5, list_location[index_data], formatone)
    sheet_now.write(index_data + 2, 6, list_feichu[index_data], formatone)

    # sheet_now.write(index_data + 2, 2, list_link[index_data], formatone)
workbook_to_write.close()
length_filedoc_list = len(list_filename_doc)
print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
tkMessageBox.showinfo('完成啦'.decode('gbk'),
                      '数据抓取完毕，共抓取到%s个结果！保存在《个人PLM上传文档数量-%s.xlsx》中'.decode('gbk') % (length_filedoc_list, timestamp))
