import pandas as pd
import numpy as np
import math
# data=pd.read_excel("C:\\Users\\Zh_uu\\Desktop\\test2\\data1.xlsx")
l=pd.read_excel("C:\\Users\\Zh_uu\\Desktop\\test2\\缺少.xlsx")
# print(l)
m=l.values.tolist()
print(m)
unique=[]
for i in m:
    unique.append(i)
unique=np.ravel(unique)
print(unique)


# l=data['unit'].tolist()
# for i in l:
#     try:
#         unit.extend(i.split(','))
#     except:
#         unit.append(i)
# #print(unit)
# unique=[]
# for ele in unit:
#     if ele not in unique:
#         unique.append(ele)
# #print(unique)
# print(unique[1])
# print(type(unique))

# import xlrd
# import os
# from xlutils.copy import copy
# from xlwt import Style
#
#
# def writeExcel(styl=Style.default_style):
#     rb = xlrd.open_workbook("company5", formatting_info=True)
#     wb = copy(rb)
#     ws = wb.get_sheet(0)
#     ws.write(styl)
#     wb.save("company5")


import time
import webbrowser
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
driver = webdriver.Edge("C:\\Users\\Zh_uu\\Desktop\\test2\\msedgedriver.exe")
#C:/Users/Zh_uu/Desktop/test2/chromedriver.exe
driver.get("https://www.qcc.com")
time.sleep(20)
product=[]

try:
    for i in range(len(unique)):
        # time.sleep(5)
        name_input=driver.find_element(by=By.CSS_SELECTOR, value='input[name=key]')
        name_input.send_keys(unique[i])
        # time.sleep(5)
        button=driver.find_element(by=By.CLASS_NAME,value="input-group-btn")
        button.click()
        time.sleep(5)
        item = {}
        try:
            tr_list=driver.find_element(by=By.XPATH, value="//div[@class='relate-info']")
            t=tr_list.find_elements(by=By.XPATH, value="./div")[0:3]
            # print(t)
            # print(len(t))
            item['公司']=unique[i]
            item['法定代表人']=t[0].find_elements(by=By.CLASS_NAME,value="val")[0].text
            item['注册资本']=t[0].find_elements(by=By.CLASS_NAME,value="val")[1].text[:-5]
            item['成立日期']=t[0].find_elements(by=By.CLASS_NAME,value="val")[2].text
            item['社会统一信用代码']=t[0].find_elements(by=By.CLASS_NAME,value="val")[3].text
            item['电话']=t[1].find_elements(by=By.CLASS_NAME,value="val")[0].text
            item['邮箱']=t[1].find_elements(by=By.CLASS_NAME,value="val")[1].text
            item['地址']=t[2].text[4:]
            # print(t[2].find_element(by=By.CLASS_NAME,value="copy-value address-map").text)
            product.append(item)
            driver.back()
        except:
            item['公司'] = unique[i]
            item['法定代表人']=''
            item['注册资本']=''
            item['成立日期']=''
            item['社会统一信用代码']=''
            item['电话']=''
            item['邮箱']=''
            item['地址']=''
            # print(t[2].find_element(by=By.CLASS_NAME,value="copy-value address-map").text)
            product.append(item)
            driver.back()
except:
    colname=['公司','法定代表人','注册资本','成立日期','社会统一信用代码','电话','邮箱','地址']
    df=pd.DataFrame(product,columns=colname)
    df.head()
    df.to_excel('./company6.xlsx')


colname=['公司','法定代表人','注册资本','成立日期','社会统一信用代码','电话','邮箱','地址']
df=pd.DataFrame(product,columns=colname)
df.head()
df.to_excel('./company缺少.xlsx')

#
# col=['名字']
# x=pd.DataFrame(data=unique,columns=col)
# print(x)
# x.to_excel('xxxx.xlsx')




