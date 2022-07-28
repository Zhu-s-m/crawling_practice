import webbrowser

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
driver = webdriver.Edge("C:\\Users\\Zh_uu\\Desktop\\test2\\msedgedriver.exe")
#C:/Users/Zh_uu/Desktop/test2/chromedriver.exe
driver.get("https://mis.kw.beijing.gov.cn/out/typt/staticHTML/infomationGkLxPiciList_1_d41d8cd98f00b204e9800998ecf8427e_25035547233.html")

list=[]
for link in driver.find_elements(by=By.XPATH,value="//div[@class='info-content']/a"):
    list.append(link)
    #list.append(link.get_attribute('href'))
#print(list)
#list[2].click()
product=[]
num=[]

for i in range(3):
        list = []
        for link in driver.find_elements(by=By.XPATH, value="//div[@class='info-content']/a"):
            list.append(link)
        # list.append(link.get_attribute('href'))
        list[i].click()


        title=[]
        for i in range(len(driver.find_elements(by=By.ID, value="content"))):
            title.append(driver.find_elements(by=By.ID, value="content")[i].text[:-6])
        title = [i for i in title if i][1:]
        print(title)



        tr_list = driver.find_elements(by=By.XPATH,value="//table[@class='text-word']/tbody/tr")[1:]
        # print(tr_list)
        if len(tr_list):
            j=0
            for tr in tr_list:
                item = {}
                try:
                     item['id']=tr.find_elements(by=By.XPATH,value="./td[1]")[0].text
                     item['pro_name']=tr.find_elements(by=By.XPATH,value="./td[2]")[0].text
                     item['top_name'] = tr.find_elements(by=By.XPATH, value="./td[3]")[0].text
                     item['unit'] = tr.find_elements(by=By.XPATH, value="./td[4]")[0].text
                     item['person'] = tr.find_elements(by=By.XPATH, value="./td[5]")[0].text
                     item['fund'] = tr.find_elements(by=By.XPATH, value="./td[6]")[0].get_attribute("textContent")
                     item['period'] = tr.find_elements(by=By.XPATH, value="./td[7]")[0].get_attribute("textContent")
                     item['class']=title[j]
                     #print(tr.find_elements(by=By.XPATH, value="./td[6]")[0].is_displayed())
                     product.append(item)
                except:
                    print('next')
                    j=j+1
        driver.back()
print(product)

import pandas as pd
colname=['id','pro_name','top_name','unit','person','fund','period','class']
df=pd.DataFrame(product,columns=colname)
df.head()
df.to_excel('./data2.xlsx')