import pandas as pd
import numpy as np
import math

# l=pd.read_excel("C:\\Users\\Zh_uu\\Desktop\\test2\\333.xlsx")
# # print(l)
# m=l.values.tolist()
# unique=[]
# for i in m:
#     unique.append(i)
# unique=np.ravel(unique)
# print(unique)
# list=[]
# for i in range(len(unique)):
#     unique[i]=unique[i][9:]
# print(unique)
# colname=['end_date']
# df=pd.DataFrame(unique,columns=colname)
# df.head()
# df.to_excel('./end_date.xlsx')

l=pd.read_excel("C:\\Users\\Zh_uu\\Desktop\\test2\\relate.xlsx")
#print(l['company'])
for i in range(len(l['company'])):
    try:
        #print(l['company'][i])
        l['company'][i]=l['company'][i].split(',')
    except:
        l['company'][i] = l['company'][i]
m=l['company'].tolist()
#print(m)
#print(l['top_name'])

#print(len(l['top_name']))
product=[]
# for i in range(len(l['top_name'])):
for i in range(len(l['top_name'])):
    #print(np.size(m[i]))
    #print(m[i])
    for j in range(np.size(m[i])):
        item={}
        item['topic']=l['top_name'][i]
        #print(l['top_name'][i])


        item['company']=m[i][j]
        #print(m[i],'xx')
        #print(m[i][j])
        #print(i)
        product.append(item)
#
print(product)
colname=['topic','company']
df=pd.DataFrame(product,columns=colname)
df.head()
df.to_excel('./rr.xlsx')
#print(m[917])
