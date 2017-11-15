# -*- coding:utf8 -*-

import time
import os
import xlwt
import numpy as np
import xlsxwriter
import pandas as pd
print os.getcwd()
from selenium import webdriver
driver = webdriver.Chrome()

data=pd.read_excel('*******')
print data.head()

driver.get('*********')

time.sleep(5)

driver.find_element_by_id("J-input-user").send_keys("***")

time.sleep(2)
driver.find_element_by_id("password_rsainput").send_keys("***")

time.sleep(5)
driver.find_element_by_id("J-login-btn").click()

time.sleep(5)
driver.get('********')

num= len(data)

for i in range(0,num):

    time.sleep(3)
    driver.find_element_by_name("userName").send_keys(data["xm"][i])
    time.sleep(2)
    driver.find_element_by_name("certNo").send_keys(data["shfzh18"][i])

    time.sleep(1)
    driver.find_element_by_id("submit").click()

    time.sleep(2)
    t1=driver.find_element_by_xpath("//*[@id='container']/div/form/div[2]").text
    print t1
    data["shfzh18"][i]=t1
    time.sleep(3)
    driver.get('*******')
    print data["shfzh18"].head(i)


data1 = data[["ajbh","shfzh18"]]

num1 = data1.shape[0]
num2 = data1.shape[1]


workbook = xlsxwriter.Workbook('dataresult.xlsx') # 建立文件

worksheet = workbook.add_worksheet() # 建立sheet

for i in xrange(2):
    if i==0:
        for j in xrange(num1):
            worksheet.write(j,i,data1["ajbh"][j])      #把获取到的值写入文件对应的行列
    else:
        for j in xrange(num1):
            worksheet.write(j,i,data1["shfzh18"][j])
workbook.close()