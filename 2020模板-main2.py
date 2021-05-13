# -*- coding: utf-8 -*-

#!/usr/bin/env python
# coding: utf-8

# In[导入外部库]
# Import some modules
from copy import deepcopy
from pandas import DataFrame
import pandas as pd
import numpy as np
import re
import os
import PIL.Image as Image
import base64
import math


from collections import Counter
from datetime import datetime
from datetime import datetime as dt

from pyecharts.charts import Bar,Pie,Tab,Line,Grid,Timeline
from pyecharts.components import Table
from pyecharts import options as opts
from pyecharts.options import ComponentTitleOpts
from pyecharts.commons import utils
from pyecharts.components import Image as echart_image
from pyecharts.globals import ThemeType

#from pypinyin import lazy_pinyin
lazy_pinyin = lambda x:x


# In[用于求物料分析中的合计功能] 
#用于求物料分析中的合计功能
def SUM(dataframe, headers):
    result = ['合计']
    for i in headers[1:-1]:
        temp = sum(dataframe[i].fillna(0).apply(lambda x:float(x)))
        result.append(temp)
    #完成率不进行求和
    result.append('%.2f%%' % (result[-2]*100/result[-1]))
    return result


# In[增量分析----数据处理]
PATH = os.getcwd()
df = pd.read_excel(PATH+'/2019项目数据.xlsx')
df2 = pd.read_excel(PATH+'/2018春季项目.xls')
project_2020_df = pd.read_excel(PATH+'/2020项目数据.xlsx',usecols=['项目','下达日期'])
project_2020_df = project_2020_df[project_2020_df['下达日期'].dt.date>datetime(2019,12,31).date()]



df1 = df[['项目','下达日期']]
df2 = df2[['项目','下达日期']]


# 按周统计2018和2019两年的项目数增量#
week_2019 = []
week_2018 = []
week_2020 = []
date_list_2019 = list(df1.下达日期)
date_list_2018 = list(df2.下达日期)
date_list_2020 = list(project_2020_df.下达日期)
for i in range(len(date_list_2019)):
    week_2019.append(date_list_2019[i].isocalendar()[1])
for i in range(len(date_list_2018)):
    week_2018.append(date_list_2018[i].isocalendar()[1])
for i in range(len(date_list_2020)):
    week_2020.append(date_list_2020[i].isocalendar()[1])
count_dict_2019 = dict(Counter(week_2019))
count_dict_2018 = dict(Counter(week_2018))
count_dict_2020 = dict(Counter(week_2020))
dic_2019 = dict.fromkeys(range(1,53),0)
dic_2018 = dict.fromkeys(range(1,53),0)
dic_2020 = dict.fromkeys(range(1,53),0)
for k,v in count_dict_2019.items():
    dic_2019[k]=count_dict_2019[k]
for k,v in count_dict_2018.items():
    dic_2018[k]=count_dict_2018[k]
for k,v in count_dict_2020.items():
    dic_2020[k]=count_dict_2020[k]


# 按周统计2018和2019两年的提案项目增量#

#COLs=['物料名称','下单日期','数量（个）']
COLs=['状态','项目名称','物料名称','DM','主视觉设计师','下单日期','数量（个）']
df_2019 = pd.read_excel(PATH+'/2019物料数据.xlsx',usecols=COLs)
df_2018 = pd.read_excel(PATH+'/2018物料数据.xls',usecols=COLs)
df_2020 = pd.read_excel(PATH+'/2020物料数据.xlsx',usecols=COLs)
df_2020['下单日期']= pd.to_datetime(df_2020.下单日期,format =('%Y-%m-%d')).dt.floor('d')
df_2020_2 = df_2020[df_2020['下单日期'].dt.date>datetime(2019,12,31).date()]

df_2019['下单日期']= pd.to_datetime(df_2019.下单日期,format =('%Y-%m-%d')).dt.floor('d')
df_2018['下单日期']= pd.to_datetime(df_2018.下单日期,format =('%Y-%m-%d')).dt.floor('d')
#df_2020['下单日期']= pd.to_datetime(df_2020.下单日期,format =('%Y-%m-%d')).dt.floor('d')

proposed_material_2019=df_2019[df_2019.物料名称=='提案物料']
proposed_material_2018=df_2018[df_2018.物料名称=='提案物料']
proposed_material_2020=df_2020_2[df_2020_2.物料名称=='提案物料']
signed_material_2019=df_2019[df_2019.物料名称!='提案物料']
signed_material_2018=df_2018[df_2018.物料名称!='提案物料']
signed_material_2020=df_2020_2[df_2020_2.物料名称!='提案物料']


proposed_material_2020_week = []
proposed_material_2019_week = []
proposed_material_2018_week = []
proposed_material_2020_date = list(proposed_material_2020.下单日期)
proposed_material_2019_date = list(proposed_material_2019.下单日期)
proposed_material_2018_date = list(proposed_material_2018.下单日期)
for i in range(len(proposed_material_2020_date)):
    proposed_material_2020_week.append(proposed_material_2020_date[i].isocalendar()[1])
for i in range(len(proposed_material_2019_date)):
    proposed_material_2019_week.append(proposed_material_2019_date[i].isocalendar()[1])
for i in range(len(proposed_material_2018_date)):
    proposed_material_2018_week.append(proposed_material_2018_date[i].isocalendar()[1])
count_proposed_material_2020 = dict(Counter(proposed_material_2020_week))
count_proposed_material_2019 = dict(Counter(proposed_material_2019_week))
count_proposed_material_2018 = dict(Counter(proposed_material_2018_week))
dict_proposed_material_2020 = dict.fromkeys(range(1,53),0)
dict_proposed_material_2019 = dict.fromkeys(range(1,53),0)
dict_proposed_material_2018 = dict.fromkeys(range(1,53),0)
for k,v in count_proposed_material_2020.items():
    dict_proposed_material_2020[k]=count_proposed_material_2020[k]
for k,v in count_proposed_material_2019.items():
    dict_proposed_material_2019[k]=count_proposed_material_2019[k]
for k,v in count_proposed_material_2018.items():
    dict_proposed_material_2018[k]=count_proposed_material_2018[k]


# 按周统计2018和2019两年的签单物料数增量#

signed_material_2020_week = []
signed_material_2019_week = []
signed_material_2018_week = []
signed_material_2020_date = list(signed_material_2020.下单日期)
signed_material_2019_date = list(signed_material_2019.下单日期)
signed_material_2018_date = list(signed_material_2018.下单日期)
for i in range(len(signed_material_2020_date)):
    signed_material_2020_week.append(signed_material_2020_date[i].isocalendar()[1])
for i in range(len(signed_material_2019_date)):
    signed_material_2019_week.append(signed_material_2019_date[i].isocalendar()[1])
for i in range(len(signed_material_2018_date)):
    signed_material_2018_week.append(signed_material_2018_date[i].isocalendar()[1])
signed_material_2020.loc[:,'周数']=signed_material_2020_week
signed_material_2019.loc[:,'周数']=signed_material_2019_week
signed_material_2018.loc[:,'周数']=signed_material_2018_week
count_signed_material_2020 =dict(signed_material_2020.groupby('周数')['数量（个）'].apply(sum))
count_signed_material_2019 =dict(signed_material_2019.groupby('周数')['数量（个）'].apply(sum))
count_signed_material_2018 =dict(signed_material_2018.groupby('周数')['数量（个）'].apply(sum))
dict_signed_material_2020 = dict.fromkeys(range(1,53),0)
dict_signed_material_2019 = dict.fromkeys(range(1,53),0)
dict_signed_material_2018 = dict.fromkeys(range(1,53),0)
for k,v in count_signed_material_2020.items():
    dict_signed_material_2020[k]=count_signed_material_2020[k]
for k,v in count_signed_material_2019.items():
    dict_signed_material_2019[k]=count_signed_material_2019[k]
for k,v in count_signed_material_2018.items():
    dict_signed_material_2018[k]=count_signed_material_2018[k]


# In[求时间序列的函数]
# 求时间序列的函数

def get_date_list(begin_date,end_date):
    date_list = [x.strftime('%Y-%m-%d') for x in list(pd.date_range(start=begin_date, end=end_date))]
    return date_list


# In[存量分析----数据处理]
# 存量分析


items_inventory = pd.read_excel(PATH+'/项目和物料状态表.xlsx',sheet_name='项目状态表')
material_inventory = pd.read_excel(PATH+'/项目和物料状态表.xlsx',sheet_name='物料状态表')
items_inventory.drop(items_inventory.index[0], inplace=True)
items_inventory.rename({"Unnamed: 0":"项目编号","Unnamed: 1":"项目"}, axis="columns", inplace=True)
material_inventory.drop(material_inventory.index[0],inplace=True)
material_inventory.rename({"Unnamed: 0":"项目编号","Unnamed: 1":"项目","Unnamed: 2":"物料"}, 
                          axis="columns", inplace=True)

items_inventory=items_inventory.drop(['项目编号','项目'],axis=1)

proposal=material_inventory[material_inventory.物料=='提案物料']
proposal=proposal.drop(['项目编号','项目','物料'],axis=1)

dateList =[i for i in get_date_list('2020-01-01','2020-09-01')]


#删除全为空值的行数据
items_inventory.columns=dateList
items_inventory=items_inventory.dropna(axis=0,how='all').loc[:,:'2020-06-30'] 

proposal.columns=dateList
proposal=proposal.dropna(axis=0,how='all').loc[:,:'2020-06-30'] 

        
#统计每一列每种状态的项目数
i_col_count=[]
#列出所有可能的取值
names = [np.nan, '取消', '进行中', '待确稿', '未分配']
for i in items_inventory.columns:
    #用默认值为0，键为names构造一个映射
    items_orignal = dict.fromkeys(names,0)
    items_orignal.update(dict(items_inventory[i].value_counts()))
    i_col_count.append(items_orignal)

  
i_col1_count=[]

#列出所有可能的取值
#names = [np.nan, '取消', '进行中', '待确稿', '未分配']
for i in proposal.columns:
    #用默认值为0，键为names构造一个映射
    proposal_orignal = dict.fromkeys(names,0)
    proposal_orignal.update(dict(proposal[i].value_counts()))
    i_col1_count.append( proposal_orignal)
    
    
ongoing=[]                     #进行中项目
unallocated=[]                 #未分配项目
proposal_ongoing=[]            #进行中提案项目
proposal_unallocated=[]        #未分配提案项目
signed_ongoing=[]              #进行中签单项目
signed_unallocated=[]          #未分配签单项目

for i in i_col_count:
    ongoing.append(i['进行中'])
    unallocated.append(i['未分配'])

for i in i_col1_count:
    proposal_ongoing.append(i['进行中'])
    proposal_unallocated.append(i['未分配'])

signed_ongoing=[ongoing[i]-proposal_ongoing[i] for i in range(len(ongoing))]
signed_unallocated=[unallocated[i]-proposal_unallocated[i] for i in range(len(unallocated))]


# In[总量分析----数据处理]

# 项目数总量关键点

## 合并2015-2020的项目数据
file1=PATH+'/2015春季项目.xls'
file2=PATH+'/2016春季项目.xls'
file3=PATH+'/2017春季项目.xls'
file4=PATH+'/2018春季项目.xls'
file5=PATH+'/2019项目数据.xlsx'
file6 = PATH + '/2020项目数据.xlsx'
file=[file1,file2,file3,file4,file5, file6]
files=[]
cols=['项目','下达日期']
for i in file:
    files.append(pd.read_excel(i,usecols=cols,encoding = 'utf-8'))
writer = pd.ExcelWriter(PATH+'/2015-2020春季项目.xlsx')
pd.concat(files).to_excel(writer,'Sheet1',index=False)
writer.save()


# 项目数关键时间点
df8=pd.read_excel(PATH+'/2015-2020春季项目.xlsx')
df9 = df8.copy()
df9['项目数']=1
df9['年']=df9['下达日期'].dt.year
df9['月']=df9['下达日期'].dt.month
df9['日']=df9['下达日期'].dt.day
month=[]
day=[]
for i in df9['月']:
    month.append(str(i).zfill(2))
for i in df9['日']:
    day.append(str(i).zfill(2))
df9['月']=month
df9['日']=day
df9['月/日']=df9['月'].map(str)+'-'+df9['日'].map(str)
todayItemsSum=df9[df9.年==2020].项目数.sum()
df9['下达日期']= pd.to_datetime(df9.下达日期,format =('%y-%m-%d')).dt.floor('d')
df9 = df9.sort_values(by = '下达日期')
years=list(df9.年.drop_duplicates())
yearNum=len(years)

#将每一年的日期限定在6.30之前
result_data = []
for year in years:
    temp = df9[df9['年'] == year]
    result_data.append(temp[temp['月'].map(int).isin(range(1,7))])

df9 = pd.concat(result_data,ignore_index=True)

del result_data


#选取关键点策略函数,如果跳过了某一个关键点则选择稍比其大的点来代替
def modi_keypoint(data, key_points, end):
    """
    param : data 数据
    param : key_points 关键点列表
    param : end 数据的最大值
    return : 关键点的日期与数据
    """
    item = []
    data_temp = data.sort_values(by='sum_items').reset_index(drop=True)
    index  = 0
    for key in  sorted(key_points):
        if key > end:
            return item
        for index_item in range(index, len(data_temp)):
            if data_temp['sum_items'][index_item] == key:
                item.append(key)
                index = index_item
                break
    
            elif data_temp['sum_items'][index_item] > key:
                item.append(data_temp['sum_items'][index_item])
                index = index_item
                break
                
    return item

#进行选取关键点
def findItemKeyPoint():
    totalItems=[]
    item=[]
    for i in range(yearNum):
        totalItems.append(df9[df9.年==years[i]].项目数.sum())
        itemPoint=list(range(0, 550,50))+totalItems
        df9['sum_items']=df9[df9.年==years[i]]['项目数'].cumsum()
        itemPoint = modi_keypoint(df9, itemPoint, totalItems[-1])
        raw_data = df9[df9['sum_items'].isin(itemPoint)]
#        两种方式进行跟新
        raw_data['sum_items'][:-1][(raw_data['sum_items']%100).astype(bool)] = \
        raw_data['sum_items'][:-1][(raw_data['sum_items']%100).astype(bool)]//100*100
#        raw_data['sum_items'] = raw_data['sum_items'].apply(lambda x: x//100*100 if (x%100 and x!= totalItems[-1]) else x)
        item.append(raw_data)
        totalItems=[]
    return item

def concatList():
    result = findItemKeyPoint()
    itemsYear = pd.concat(result,axis=0,ignore_index=True)
    return itemsYear

def getFullYearData():
    result = concatList()
    full_years_map = {}
    for year in  result['年'].unique():
        data = result[result['年']==year][['月/日', 'sum_items']].set_index('月/日').to_dict()['sum_items']
        full_years_map[year] = data
    return  full_years_map


Date_List =[i[-5:] for i in get_date_list('2020-01-01','2020-06-30')]
Date_List_length=len(Date_List)
date_2015 = dict(zip(Date_List,[np.nan for i in range(Date_List_length)]))
date_2016 = dict(zip(Date_List,[np.nan for i in range(Date_List_length)]))
date_2017 = dict(zip(Date_List,[np.nan for i in range(Date_List_length)]))
date_2018 = dict(zip(Date_List,[np.nan for i in range(Date_List_length)]))
date_2019 = dict(zip(Date_List,[np.nan for i in range(Date_List_length)]))
date_2020 = dict(zip(Date_List,[np.nan for i in range(Date_List_length)]))
#得到所有年份项目数的数据
FullYearData = getFullYearData()
date_2015.update(FullYearData[2015])
date_2016.update(FullYearData[2016])
date_2017.update(FullYearData[2017])
date_2018.update(FullYearData[2018])
date_2019.update(FullYearData[2019])
date_2020.update(FullYearData[2020])
# 物料数总量关键点

# 合并2015-2020的物料数据
File1=PATH+'/2015物料数据.xls'
File2=PATH+'/2016物料数据.xls'
File3=PATH+'/2017物料数据.xls'
File4=PATH+'/2018物料数据.xls'
File5=PATH+'/2019物料数据.xlsx'
File6=PATH+'/2020物料数据.xlsx'
File=[File1,File2,File3,File4,File5, File6]
Files=[]
cols=['项目名称','下单日期','数量（个）']
for i in File:
    Files.append(pd.read_excel(i,usecols=cols,encoding = 'utf-8'))
writer = pd.ExcelWriter(PATH+'/2015-2020物料数据.xlsx')
pd.concat(Files).to_excel(writer,'Sheet1',index=False)
writer.save()

###############################################################
#手动调整2018-2017的数量个数
#df = pd.read_excel(PATH+'/2015-2020物料数据.xlsx',encoding='utf8')
#for i in range(len(df)):
#    if df.ix[i,'下单日期'].startswith(('2017','2018')):
#        df.ix[i,'数量（个）'] = 1.5
#writer = pd.ExcelWriter(PATH+'/2015-2020物料数据.xlsx')
#df.to_excel(writer,'Sheet1',index=False)
#writer.save()
    
###############################################################   

## 物料数关键时间点
material_df=pd.read_excel(PATH+'/2015-2020物料数据.xlsx')
material_df['下单日期']=pd.to_datetime(material_df['下单日期'],format=('%Y-%m-%d')).dt.floor('d')
material_df['年']=material_df['下单日期'].dt.year
material_df['月']=material_df['下单日期'].dt.month
material_df['日']=material_df['下单日期'].dt.day
month=[]
day=[]
for i in material_df['月']:
    month.append(str(i).zfill(2))
for i in material_df['日']:
    day.append(str(i).zfill(2))
material_df['月']=month
material_df['日']=day
material_df['月/日']=material_df['月'].map(str)+'-'+material_df['日'].map(str)
todayMaterialSum=material_df[material_df.年==2020]['数量（个）'].sum()
material_df = material_df.sort_values(by = '下单日期')
years=list(material_df.年.drop_duplicates())
yearNum=len(years)


#将每一年的日期限定在6.30之前
result_data = []
for year in years:
    temp = material_df[material_df['年'] == year]
    result_data.append(temp[temp['月'].map(int).isin(range(1,7))])

material_df = pd.concat(result_data,ignore_index=True)

del result_data





def findMaterialKeyPoint():
    totalMaterials=[]
    material=[]
    for i in range(yearNum):
        totalMaterials.append(material_df[material_df.年==years[i]]['数量（个）'].sum())
        materialPoint= list(range(0, 5000,400))+totalMaterials
        material_df['sum_items']=material_df[material_df.年==years[i]]['数量（个）'].cumsum()
        materialPoint = modi_keypoint(material_df, materialPoint, totalMaterials[-1])    
        raw_data = material_df[material_df['sum_items'].isin(materialPoint)]
        raw_data['sum_items'][:-1][(raw_data['sum_items']%1000).astype(bool)] = \
        raw_data['sum_items'][:-1][(raw_data['sum_items']%1000).astype(bool)]//1000*1000
#        raw_data['sum_items'] = raw_data['sum_items'].apply(lambda x: x if (not x%1000 or x == totalMaterials[-1]) else x//1000*1000)
        material.append(raw_data)
        totalMaterials=[]
    return material


def concatMaterialList():
    result = findMaterialKeyPoint()
    materialsPerYear = pd.concat(result,axis=0,ignore_index=True)
    return materialsPerYear



def getFullYearMaterialData():
    result = concatMaterialList()
    full_years_map = {}
    for year in  result['年'].unique():
        data = result[result['年']==year][['月/日', 'sum_items']].set_index('月/日').to_dict()['sum_items']
        full_years_map[year] = data
    return  full_years_map

Date_List =[i[-5:] for i in get_date_list('2020-01-01','2020-06-30')]
Date_List_length=len(Date_List)
Material_2015 = dict(zip(Date_List,[np.nan for i in range(Date_List_length)]))
Material_2016 = dict(zip(Date_List,[np.nan for i in range(Date_List_length)]))
Material_2017 = dict(zip(Date_List,[np.nan for i in range(Date_List_length)]))
Material_2018 = dict(zip(Date_List,[np.nan for i in range(Date_List_length)]))
Material_2019 = dict(zip(Date_List,[np.nan for i in range(Date_List_length)]))
Material_2020 = dict(zip(Date_List,[np.nan for i in range(Date_List_length)]))
#得到所有年份的物料关键点数据
FullYearMaterialData = getFullYearMaterialData()

Material_2015.update(FullYearMaterialData[2015])
Material_2016.update(FullYearMaterialData[2016])
Material_2017.update(FullYearMaterialData[2017])
Material_2018.update(FullYearMaterialData[2018])
Material_2019.update(FullYearMaterialData[2019])
Material_2020.update(FullYearMaterialData[2020])
## 完稿画面数关键时间点


finished_picture_df=pd.read_excel(PATH+'/2020物料数据.xlsx')
finished_picture_df['下单日期']=pd.to_datetime(finished_picture_df['下单日期'],format=('%Y-%m-%d')).dt.floor('d')
finished_picture_df['年']=finished_picture_df['下单日期'].dt.year
finished_picture_df['月']=finished_picture_df['下单日期'].dt.month
finished_picture_df['日']=finished_picture_df['下单日期'].dt.day
month=[]
day=[]
for i in finished_picture_df['月']:
    month.append(str(i).zfill(2))
for i in finished_picture_df['日']:
    day.append(str(i).zfill(2))
finished_picture_df['月']=month
finished_picture_df['日']=day
finished_picture_df['月/日']=finished_picture_df['月'].map(str)+'-'+finished_picture_df['日'].map(str)
todayItemsSum=finished_picture_df[finished_picture_df.年==2020].完稿画面数.sum()
finished_picture_df['下单日期']= pd.to_datetime(finished_picture_df.下单日期,format =('%y-%m-%d')).dt.floor('d')
finished_picture_df = finished_picture_df.sort_values(by = '下单日期')

totalItems=[]
item=[]
def find_KeyPoint():
    totalItems.append(finished_picture_df[finished_picture_df.年==2020].完稿画面数.sum())
    itemPoint=list(range(0,6000,500))+totalItems
    finished_picture_df['sum_items']=finished_picture_df[finished_picture_df.年==2020]['完稿画面数'].cumsum()
    itemPoint = modi_keypoint(finished_picture_df, itemPoint, totalItems[-1])    
    raw_data = finished_picture_df[finished_picture_df['sum_items'].isin(itemPoint)]
    raw_data['sum_items'][:-1][(raw_data['sum_items']%5000).astype(bool)] = \
    raw_data['sum_items'][:-1][(raw_data['sum_items']%5000).astype(bool)]//5000*5000
#    raw_data['sum_items'] = raw_data['sum_items'].apply(lambda x: x if (not x%5000 or x == totalItems[-1]) else x//5000*5000)
    item.append(raw_data)
    return item

def getData_2020(year):
    data=dict(zip(find_KeyPoint()[0]['月/日'],find_KeyPoint()[0]['sum_items']))
    return  data


Date_List =[i[-5:] for i in get_date_list('2020-01-01','2020-06-30')]
Date_List_length=len(Date_List)
finished_pictures_2019 = dict(zip(Date_List,[np.nan for i in range(Date_List_length)]))
finished_pictures_2019.update(getData_2020(2020))


# In[AE分析----数据处理]

## AE签单项目分析
names=['赵俊芳','唐祎杰','朱颖','胡芳怡','刘然','黄宇澄','李忠璐','孙立飞',
       '姜渊韬','王毅哲','孙金迅','黄婉露','谢诗琪','黄颖倩','夏澋婧','赵中元',
       '叶富军','潘琛','刘波','李家星','季金凯','黄杨','方慧','王佳城','田明','郑海港']

categories=['总量','进行中','已完成','未完成']
categoryNum = len(categories)

df2020 = pd.read_excel(PATH+'/2020项目数据.xlsx')

df3 = df2020[['状态','DM','项目']]
df3=df3[df3['状态'].isin(['进行中','已完成','未完成'])]
df3=df3[df3['DM'].isin(names)]
items_dm = df3.groupby('DM')['项目'].count()
itemsNumList = items_dm.values
new_name = items_dm.index.to_list()
namesNum = len(new_name)
itemsAll = pd.Series(itemsNumList,index=[['总量'for i in range(namesNum)],new_name])
itemsByStates = df3.groupby(['状态','DM'])['项目'].count()
items_df = pd.concat([itemsAll,itemsByStates]).to_frame() 
items_df.reset_index(inplace=True)
items_df.columns=['状态','DM','项目量']
items_df = pd.concat([items_df,pd.DataFrame([[np.nan,np.nan, np.nan]],
                                index=items_df.index, 
                                columns=['进行中','已完成','未完成'])], axis=1)
underway=[]           #进行中
completed=[]          #已完成
unfinished=[]         #未完成
DMs = []              #设计师
categories_df = []
for i in range(categoryNum):
    data = items_df[items_df['状态'] == categories[i]].sort_values(by="项目量",ascending=False)
    data = data.reset_index(drop=True)
    DMs.append(data.DM)
    categories_df.append(data)
for i in range(len(DMs[0])):    
    underway.append(int(df3[(df3['DM']==DMs[0][i]) & (df3['状态'] == categories[1])].项目.count()))
    completed.append(int(df3[(df3['DM']==DMs[0][i]) & (df3['状态'] == categories[2])].项目.count()))
    unfinished.append(int(df3[(df3['DM']==DMs[0][i]) & (df3['状态'] == categories[3])].项目.count()))
categories_df[0].loc[:,'进行中']=underway
categories_df[0].loc[:,'已完成']= completed
categories_df[0].loc[:,'未完成']= unfinished

## AE签单物料分析
cols = ['状态','项目编号','项目名称','DM','完稿画面数','数量（个）']
df5 = pd.read_excel(PATH+'/2020物料数据.xlsx',usecols=cols)

material_df5 = df5.copy()
material_df5=material_df5[material_df5['状态'].isin(['进行中','已完成','未完成'])]
#names=['赵俊芳','唐祎杰','朱颖','胡芳怡','刘然','黄宇澄','李忠璐','孙立飞',
#       '姜渊韬','王毅哲','孙金迅','黄婉露','谢诗琪','黄颖倩','夏澋婧','赵中元',
#       '叶富军','潘琛','刘波','李家星','季金凯','黄杨','方慧','王佳城','田明','郑海港']
material_df5=material_df5[material_df5['DM'].isin(names)]
materials_dm = material_df5.groupby('DM')['数量（个）'].apply(sum)
materialsNumList = materials_dm.values
new_names = materials_dm.index.to_list()
namesNum = len(new_names)
materialsAll = pd.Series(materialsNumList,index=[['总量'for i in range(namesNum)],new_names])
materialsByStates = material_df5.groupby(['状态','DM'])['数量（个）'].apply(sum)
materials_df5 = pd.concat([materialsAll,materialsByStates]).to_frame() 
materials_df5.reset_index(inplace=True)
materials_df5.columns=['状态','DM','项目量']
materials_df5 = pd.concat([materials_df5,pd.DataFrame([[np.nan,np.nan, np.nan]],
                                index=materials_df5.index, 
                                columns=['进行中','已完成','未完成'])], axis=1)
state_class=['总量','进行中','已完成','未完成']
stateNum = len(state_class)
material_underway=[]           #进行中签单物料
material_completed=[]          #已完成签单物料
material_unfinished=[]         #未完成签单物料
material_DMs = []              #签单物料设计师
categories_df5 = []
for i in range(stateNum):
    data1 = materials_df5[materials_df5['状态'] == state_class[i]].sort_values(by="项目量",ascending=False)
    data1 = data1.reset_index(drop=True)
    material_DMs.append(data1.DM)
    categories_df5.append(data1)

for i in range(len(material_DMs[0])):    
    material_underway.append(int(material_df5[(material_df5['DM']==material_DMs[0][i]) & 
                                     (material_df5['状态'] == state_class[1])]['数量（个）'].sum()))
    material_completed.append(int(material_df5[(material_df5['DM']==material_DMs[0][i]) & 
                                      (material_df5['状态'] == state_class[2])]['数量（个）'].sum()))
    material_unfinished.append(int(material_df5[(material_df5['DM']==material_DMs[0][i]) & 
                                       (material_df5['状态'] == state_class[3])]['数量（个）'].sum()))
categories_df5[0].loc[:,'进行中']=material_underway
categories_df5[0].loc[:,'已完成']= material_completed
categories_df5[0].loc[:,'未完成']= material_unfinished


## AE完稿画面数分析
finished_picture_df7 = df5.copy()
finished_picture_df7=finished_picture_df7[finished_picture_df7['状态'].isin(['进行中','已完成','未完成'])]
#dms=['赵俊芳','唐祎杰','朱颖','胡芳怡','刘然','黄宇澄','李忠璐','孙立飞',
#       '姜渊韬','王毅哲','孙金迅','黄婉露','谢诗琪','黄颖倩','夏澋婧','赵中元',
#       '叶富军','潘琛','刘波','李家星','季金凯','黄杨','方慧','王佳城','田明','郑海港']
finished_picture_df7=finished_picture_df7[finished_picture_df7['DM'].isin(names)]
finished_picture_dm = finished_picture_df7.groupby('DM')['完稿画面数'].sum()
finished_pictureNumList = finished_picture_dm.values
new_dms = finished_picture_dm.index.to_list()
dmsNum = len(new_dms)
finished_pictureAll = pd.Series(finished_pictureNumList,index=[['总数'for i in range(dmsNum)],new_dms])
finished_pictureByStates = finished_picture_df7.groupby(['状态','DM'])['完稿画面数'].sum()
finished_picture_df = pd.concat([finished_pictureAll,finished_pictureByStates]).to_frame() 
finished_picture_df.reset_index(inplace=True)
finished_picture_df.columns=['状态','DM','完稿画面总数']
finished_picture_df = pd.concat([finished_picture_df,pd.DataFrame([[np.nan,np.nan, np.nan]],
                                index=finished_picture_df.index, 
                                columns=['进行中','已完成','未完成'])], axis=1)
finished_picture_categories=['总数','进行中','已完成','未完成']
finished_picture_underway = []            #进行中完稿画面
finished_picture_completed = []           #已完成完稿画面
finished_picture_unfinished = []          #未完成完稿画面
finished_picture_categories_df = []
Designers=[]
finished_pictureCategoryNum = len(finished_picture_categories)
for i in range(finished_pictureCategoryNum):
    data2 = finished_picture_df[finished_picture_df['状态'] == finished_picture_categories[i]].sort_values(by="完稿画面总数",
                                                                                ascending=False)
    data2 = data2.reset_index(drop=True)
    Designers.append(data2.DM)
    finished_picture_categories_df.append(data2)
Designers0Num = len(Designers[0])
for i in range(Designers0Num):
    finished_picture_underway.append(int(finished_picture_df7[(finished_picture_df7['DM']==Designers[0][i]) 
                            & (finished_picture_df7['状态'] == finished_picture_categories[1])].完稿画面数.sum()))
    finished_picture_completed.append(int(finished_picture_df7[(finished_picture_df7['DM']==Designers[0][i]) 
                            & (finished_picture_df7['状态'] == finished_picture_categories[2])].完稿画面数.sum()))
    finished_picture_unfinished.append(int(finished_picture_df7[(finished_picture_df7['DM']==Designers[0][i]) 
                            & (finished_picture_df7['状态'] == finished_picture_categories[3])].完稿画面数.sum()))
finished_picture_categories_df[0].loc[:,'进行中']= finished_picture_underway
finished_picture_categories_df[0].loc[:,'已完成']= finished_picture_completed
finished_picture_categories_df[0].loc[:,'未完成']= finished_picture_unfinished


#把签单项目、签单物料、完稿画面的数据添加到列表categories_df_list中，方便可视化中的时间轴联动
categories_df_list=[]
categories_df_list.append(categories_df[0])
categories_df_list.append(categories_df5[0])
categories_df_list.append(finished_picture_categories_df[0])


# In[设计师分析----数据处理]
df_designer=pd.read_excel(PATH+'/设计师数据.xlsx')
df_designer=df_designer[~df_designer['姓名'].isin(['兼职'])]
df_designer['类别']=df_designer['类别'].ffill()
#category_type = ['提案','平面','网页','组长','其他']
category_type = ['提案','平面','网页','其他']
category_type_Num = len(category_type)
designers = []    #设计师
distributed=[]    #未分配
on_going=[]       #进行中
draft=[]          #待确稿
unfinished=[]     #未完成
completed=[]      #已完成
workload=[]       #工作量
for i in range(category_type_Num):
    category_df = df_designer[df_designer['类别'] == category_type[i]].sort_values(by="合计",ascending=False)
    designers.append(category_df.姓名)
    distributed.append(category_df.未分配)
    on_going.append(category_df.进行中)
    draft.append(category_df.待确稿)
    unfinished.append(category_df.未完成)
    completed.append(category_df.已完成)
    workload.append(category_df.工作量)
    


# In[大区分析----数据处理]
##读取数据，将城市和大区左连接关联
cols=['项目','城市']
df_2019_data = pd.read_excel(PATH+'/2019项目数据.xlsx',usecols=cols)
df_2018_data = pd.read_excel(PATH+'/2018春季项目.xls',usecols=cols)
df_2020_data = pd.read_excel(PATH+'/2020项目数据.xlsx',usecols=cols)
cols1=['数量（个）','完稿画面数','城市','项目名称','物料名称','下单日期']
cols2=['数量（个）','城市','项目名称','物料名称','下单日期']
material_df_2019 = pd.read_excel(PATH+'/2019物料数据.xlsx',usecols=cols2)
material_df_2018 = pd.read_excel(PATH+'/2018物料数据.xls',usecols=cols2)
material_df_2020 = pd.read_excel(PATH+'/2020物料数据.xlsx',usecols=cols1)
regions = pd.read_excel(PATH+'/大区-城市匹配表.xlsx')
city_region_2019 = df_2019_data.merge(regions,how="left",on="城市")
city_region_2018 = df_2018_data.merge(regions,how="left",on="城市")
city_region_2020 = df_2020_data.merge(regions,how="left",on="城市")
city_region_material_2019 = material_df_2019.merge(regions,how="left",on="城市")
city_region_material_2018 = material_df_2018.merge(regions,how="left",on="城市")
city_region_material_2020 = material_df_2020.merge(regions,how="left",on="城市")
Regions=['上海区','华南区','华北区','华东区','华中区']

##分别统计2017、2018和2019年五个大区所对应的项目数、物料数和完稿画面数
item_region_2019=city_region_2019.groupby('大区')['项目'].count()
item_region_2018=city_region_2018.groupby('大区')['项目'].count()
item_region_2020=city_region_2020.groupby('大区')['项目'].count()
material_region_2019=city_region_material_2019.groupby('大区')['数量（个）'].apply(sum)
material_region_2018=city_region_material_2018.groupby('大区')['数量（个）'].apply(sum)
material_region_2020=city_region_material_2020.groupby('大区')['数量（个）'].apply(sum)
finished_picture_region_2020=dict(city_region_material_2020.groupby('大区')['完稿画面数'].apply(sum))
finished_picture_region_2018=dict(zip([k for k in Regions], 
         [np.nan,np.nan,np.nan,np.nan,np.nan]))
finished_picture_region_2019=dict(zip([k for k in Regions], 
         [np.nan,np.nan,np.nan,np.nan,np.nan]))

item_regions = DataFrame({
                        '2020':item_region_2020,
                        '2019':item_region_2019,
                        '2018':item_region_2018,
                       }).reindex(Regions)
material_regions = DataFrame({
                            '2020':material_region_2020,
                            '2019':material_region_2019,
                            '2018':material_region_2018,
                           }).reindex(Regions)
finished_picture_regions = DataFrame({
                                    '2020':finished_picture_region_2020,
                                    '2019':finished_picture_region_2019,
                                    '2018':finished_picture_region_2018,
                                   }).reindex(Regions)

five_regions_df=[]
five_regions_df.append(item_regions)
five_regions_df.append(material_regions)
five_regions_df.append(finished_picture_regions)

##2019年五大区的项目数，从大到小排列
items_region=dict(city_region_2020.groupby('大区')['项目'].count())
items_region_list= sorted(items_region.items(),key=lambda x:x[1],reverse=True)
items_region_key=[]
items_region_value=[]
for i in range(len(items_region_list)):
    items_region_key.append(items_region_list[i][0])
    items_region_value.append(items_region_list[i][1])
    

# In[城市分析----数据处理]
## 分别统计2017、2018和2019年各个城市所对应的项目数、物料数和完稿画面数，并从大到小排列
## 所有结果的城市均按照2018年项目数最多的10个城市来
city_2019_item=df_2019_data['城市'].value_counts()
city_2018_item=df_2018_data['城市'].value_counts()[:10]
city_2020_item=df_2020_data['城市'].value_counts()
cities=city_2018_item.index.to_list()
top10_cities_2018=list(city_2018_item.values)
top10_cities_2020=[]
top10_cities_2019=[]
for c in cities:
    top10_cities_2020.append(city_2020_item.get(c,np.nan))
    top10_cities_2019.append(city_2019_item.get(c,np.nan))

material_project_2018=dict(material_df_2018.groupby('城市')['数量（个）'].apply(sum).sort_values(ascending=False))
material_project_2019=dict(material_df_2019.groupby('城市')['数量（个）'].apply(sum).sort_values(ascending=False))
material_project_2020=dict(material_df_2020.groupby('城市')['数量（个）'].apply(sum).sort_values(ascending=False))

top10_cities_material_2018=[]
top10_cities_material_2020=[]
top10_cities_material_2019=[]
for c in cities:
    top10_cities_material_2020.append(material_project_2020.get(c, np.nan))
    top10_cities_material_2018.append(material_project_2018.get(c, np.nan))
    top10_cities_material_2019.append(material_project_2019.get(c, np.nan))
    
finished_picture_2020=dict(material_df_2020.groupby('城市')['完稿画面数'].apply(sum).sort_values(ascending=False))
finished_picture_2019=dict(zip([k for k in list(finished_picture_2020.keys())], 
         [np.nan,np.nan,np.nan,np.nan,np.nan,np.nan,np.nan,np.nan,np.nan,np.nan]))
finished_picture_2018=dict(zip([k for k in list(finished_picture_2020.keys())], 
         [np.nan,np.nan,np.nan,np.nan,np.nan,np.nan,np.nan,np.nan,np.nan,np.nan]))
top10_cities_finished_picture_2020=[]
for c in cities:
    top10_cities_finished_picture_2020.append(finished_picture_2020.get(c, np.nan))
    
finished_picture = {'city': [k for k in cities],
         '2020': [v for v in top10_cities_finished_picture_2020],
          '2019': [v for v in list(finished_picture_2019.values())],
          '2018': [v for v in list(finished_picture_2018.values())]}
finished_picture_df = pd.DataFrame.from_dict(finished_picture)
material_project = {'city': [k for k in cities],
         '2020': [v for v in top10_cities_material_2020],
          '2019': [v for v in top10_cities_material_2019],
          '2018': [v for v in top10_cities_material_2018]}
material_project_df = pd.DataFrame.from_dict(material_project)
project = {'city': [k for k in cities],
         '2020': [v for v in top10_cities_2020],
          '2019': [v for v in top10_cities_2019],
          '2018': [v for v in top10_cities_2018]}
project_df = pd.DataFrame.from_dict(project)

three_categories_df=[]
three_categories_df.append(project_df)
three_categories_df.append(material_project_df)
three_categories_df.append(finished_picture_df)


## 分别统计2019和2020年提案和签单项目占总项目的百分比
three_years_top10_city=pd.concat([material_df_2020,
                                 material_df_2019],axis=0,ignore_index=True,sort=True)
three_years_top10_city['year']=pd.to_datetime(three_years_top10_city.下单日期).dt.year
three_years_top10_city_proposal=three_years_top10_city[three_years_top10_city.物料名称=='提案物料']
proposal_data = three_years_top10_city_proposal.groupby(['城市','year'])['项目名称'].count().unstack().reset_index()
proposal_data_top10=proposal_data.sort_values(by=2019,ascending=False)[:10]
three_years_top10_city_signed=three_years_top10_city[three_years_top10_city.物料名称!='提案物料'].drop_duplicates('项目名称')
signed_data = three_years_top10_city_signed.groupby(['城市','year'])['项目名称'].count().unstack().reset_index()
signed_data_top10=signed_data[signed_data.城市.isin(proposal_data_top10.城市)].sort_values(by=2019,ascending=False)[:10]
proposal_signed_top10 = proposal_data_top10.merge(signed_data_top10,how="left",on="城市")
proposal_signed_top10['2019_proposal_%']=round((proposal_signed_top10['2019_x']/(proposal_signed_top10['2019_x']+proposal_signed_top10['2019_y']))*100,0)
proposal_signed_top10['2020_proposal_%']=round((proposal_signed_top10['2020_x']/(proposal_signed_top10['2020_x']+proposal_signed_top10['2020_y']))*100,0)
proposal_signed_top10['2019_signed_%']=round((proposal_signed_top10['2019_y']/(proposal_signed_top10['2019_x']+proposal_signed_top10['2019_y']))*100,0)
proposal_signed_top10['2020_signed_%']=round((proposal_signed_top10['2020_y']/(proposal_signed_top10['2020_x']+proposal_signed_top10['2020_y']))*100,0)


proposal_df_2019 = df_2019[df_2019.物料名称=='提案物料'].drop_duplicates('项目名称')
proposal_df_2020 = df_2020[df_2020.物料名称=='提案物料'].drop_duplicates('项目名称')
signed_df_2019 = df_2019[df_2019.物料名称!='提案物料'].drop_duplicates('项目名称')
signed_df_2020 = df_2020[df_2020.物料名称!='提案物料'].drop_duplicates('项目名称')


# In[产能分析----数据处理]
#总量
total_proposal_2020 = proposal_df_2020.shape[0]
total_proposal_2019 = proposal_df_2019.shape[0]
total_signed_2020 = signed_df_2020.shape[0]
total_signed_2019 = signed_df_2019.shape[0]

#占比
proposal_ratio_2020=round((total_proposal_2020/total_proposal_2019)*100,2)
signed_ratio_2020=round((total_signed_2020/total_signed_2019)*100,2)

#同比
df_2019_new = df_2019[df_2019.下单日期.dt.year==2019]
df_2019_new = df_2019_new[df_2019_new.下单日期.dt.month<=dt.today().month]
df_2019_new = df_2019_new[df_2019_new.下单日期.dt.day<=dt.today().day]
proposal_df_2019_new = df_2019_new[df_2019_new.物料名称=='提案物料'].drop_duplicates('项目名称')
signed_df_2019_new = df_2019_new[df_2019_new.物料名称!='提案物料'].drop_duplicates('项目名称')
total_proposal_df_2019_new = proposal_df_2019_new.shape[0]
total_signed_df_2019_new = signed_df_2019_new.shape[0]
try:
    proposal_year_on_year = round((total_proposal_2020/total_proposal_df_2019_new)*100,2)
except ZeroDivisionError:
    proposal_year_on_year = 0
try:   
    signed_year_on_year = round((total_signed_2020/total_signed_df_2019_new)*100,2)
except ZeroDivisionError:
    signed_year_on_year = 0
#存量
proposal_stock_2020 = proposal_df_2020[proposal_df_2020.状态.isin(['未分配','进行中','待确稿'])].shape[0]
signed_stock_2020 = signed_df_2020[signed_df_2020.状态.isin(['未分配','进行中','待确稿'])].shape[0]

#人均
proposal_undergoing = proposal_df_2020[proposal_df_2020.状态=='进行中']
signed_undergoing = signed_df_2020[signed_df_2020.状态=='进行中']
proposal_per_capita = round(proposal_undergoing.shape[0]/\
                            proposal_undergoing['主视觉设计师'].dropna().drop_duplicates().shape[0],2)
signed_per_capita = round(signed_undergoing.shape[0]/\
                            signed_undergoing['DM'].dropna().drop_duplicates().shape[0],2)

#重点产品
key_product_states=['未分配','进行中','已完成','未完成','取消']
shijie_df = df_2020[df_2020.物料名称.map(str).str.contains('事界')].drop_duplicates('项目名称')
shijie_total=shijie_df.shape[0]
shangxiankuai_df = df_2020[df_2020.物料名称.isin(['上线快'])].drop_duplicates('项目名称')
shangxiankuai_total = shangxiankuai_df.shape[0]
shijie_list = list(shijie_df [shijie_df.状态.isin(key_product_states)]\
     .groupby('状态')['项目名称'].count().reindex(key_product_states).values)
shangxiankuai_list = list(shangxiankuai_df [shangxiankuai_df.状态.isin(key_product_states)]\
     .groupby('状态')['项目名称'].count().reindex(key_product_states).values)
shijie_list.append(shijie_total)
shangxiankuai_list.append(shangxiankuai_total)
EVP=[0,0,8,0,0,0]
shijie_list.insert(0,'小程序')
shangxiankuai_list.insert(0,'上线快')
EVP.insert(0,'EVP调研')


designer_task_df=pd.read_excel(PATH+"/设计师任务明细表.xlsx")
designer_task_df1 = designer_task_df.drop([0],axis=0).fillna(0)
designer_task_df1=designer_task_df1.drop(['工种','工号','姓名','区域','岗位','工作量'],axis=1)

product_capacity = ['提案物料','创意主视觉','签单物料','子站设计','H5设计',\
'长图文','插画','三维','前端','后台','视频','PPT','策划','文案']
designer_task_df1.columns=product_capacity
designer_task_df1_T = designer_task_df1.T
designer_task_df1_T['总量'] = designer_task_df1_T.apply(lambda x: x.sum(), axis=1)
designer_task_df1_T_new = designer_task_df1_T[designer_task_df1_T.columns[-1]].to_frame()
designer_task_df1_T_new['人均']=round(designer_task_df1_T['总量']/(designer_task_df.shape[0]-1),2)
designer_task_df1_T_new['满载值'] = [3,4,3,2,2,3,2,2,2,2,2,2,3,3]
designer_task_df1_T_new['负载率'] = round(designer_task_df1_T_new['人均']/\
                                       designer_task_df1_T_new['满载值'],2)


# In[物料分析----数据处理]
product_analysis_df = df_2020.groupby(['状态','物料名称'])['数量（个）'].sum().unstack().T
product_analysis_df['总计'] = product_analysis_df.apply(lambda x: x.sum(), axis=1)
#将不存在的列放入一个列表中
disable_columns = []
#异常处理状态中是否有取消这一列
try:
    product_analysis_df['取消']
except:
    disable_columns.append('取消')
    print('2020物料数据表中状态一栏没有取消中')
    product_analysis_df['取消'] = 0
    
product_analysis_df['完成率'] = (product_analysis_df['已完成'].fillna(0)/(product_analysis_df['总计'].fillna(0)-product_analysis_df['取消'].fillna(0)))#.apply(lambda x:'%.2f%%'%(100*x))
product_analysis_df_sorted = product_analysis_df.sort_values(by='总计',ascending=False)


#对列进行补全
for column in ['未分配', '进行中','待确稿','取消','未完成','已完成','总计','完成率']:
    try:
        product_analysis_df_sorted[column]
    except:
        disable_columns.append(column)
        print('2020物料数据表中状态一栏没有%s' %column)
        product_analysis_df_sorted[column] = 0

product_analysis_df_sorted = product_analysis_df_sorted[['未分配', '进行中','待确稿','取消','未完成','已完成','总计','完成率']]
header = ['物料名称','未分配', '进行中','待确稿','取消','未完成','已完成','总计','完成率'] 
#因为数据中已经去掉了  再修改
#product_analysis_df_sorted = product_analysis_df_sorted[['未分配', '进行中','待确稿', '再修改','取消','未完成','已完成','总计','完成率']]
#header = ['物料名称','未分配', '进行中','待确稿', '再修改','取消','未完成','已完成','总计','完成率'] 
HE_JI = SUM(product_analysis_df_sorted,header)
HE_JI = [i  if i else '' for i in HE_JI ]
#删除合计中的再确认
#HE_JI[1] += HE_JI[2]
#HE_JI[2] += HE_JI[4]
#del HE_JI[4]
#删除排队中这一列
#del HE_JI[2]
product_analysis_df_sorted[disable_columns] = np.nan

product_analysis_df_sorted['完成率'] = product_analysis_df_sorted['完成率'].apply(lambda x:'%.2f%%'%(100*x)).replace('nan%', '0%')

product_analysis_df_sorted = product_analysis_df_sorted.reset_index('物料名称').fillna('')
#product_analysis_df_sorted = product_analysis_df_sorted.reset_index('物料名称').fillna('')
#header = ['物料名称','未分配','排队中','进行中','待确稿','再修改','取消','未完成','已完成','总计','完成率'] 
product_analysis_df_sorted = product_analysis_df_sorted.loc[:,header]

product_analysis_df_sorted['完成率'] = product_analysis_df_sorted['完成率'].apply(lambda x:x if float(x[:-1]) else '')
#暂时excel中没有删除排队中这一列，不过后续会在excel中删除排队中这一列，所以用try语句
try:
    product_analysis_df_sorted.未分配 = product_analysis_df_sorted[['排队中','未分配']].replace('',0).sum(axis=1).replace(0, '')
    product_analysis_df_sorted = product_analysis_df_sorted.drop(columns=['排队中'])    
except:
    pass
#暂时excel中没有删除再修改这一列，不过后续会在excel中删除再修改这一列，所以用try语句
try:
    product_analysis_df_sorted.进行中 = product_analysis_df_sorted[['再修改','进行中']].replace('',0).sum(axis=1).replace(0, '')
    product_analysis_df_sorted = product_analysis_df_sorted.drop(columns=['再修改'])    
except:
    pass



# In[函数---就地尝试将二维数组中浮点数转换成int类型] 

#==================================================

#将浮点数转换成int类型，不返回变量
def fun(rows):
    temp = rows.copy()
    for row in range(len(temp)):
        for index in range(len(temp[row])):
            try:
                rows[row][index] = int(temp[row][index])
            except ValueError:
                continue 





#数据可视化

# In[增量分析----可视化] 
def incrementCompareBar()->Bar: 
    bar = (
        Bar(init_opts = opts.InitOpts(bg_color='white'))    #0e2147
        .add_xaxis(['week{}'.format(w) for w in range(1, 27)])
        .add_yaxis("2018年",
                   [v for v in list(dic_2018.values())],
                   xaxis_index=0,
                   color = '#C23531',
                   gap = 0,
                   markline_opts = opts.MarkLineOpts(
                           data = [opts.MarkLineItem(name = "average",type_ = "average")],
                           label_opts=opts.LabelOpts(is_show=True,position = 'end')
                                                    )
                  )
        .add_yaxis("2019年",
                   [int(v) for v in list(dic_2019.values())],
                   xaxis_index=0,
                   color = '#2F4554',
                   gap = 0,
                   markline_opts = opts.MarkLineOpts(
                           data = [opts.MarkLineItem(name = "average",type_ = "average")],
                           label_opts=opts.LabelOpts(is_show=True,position = 'end')
                                                    )
                  )
        .add_yaxis("2020年",
                   [v for v in list(dic_2020.values())],
                   xaxis_index=0,
                   color = '#61A0A8',
                   gap = 0,
                   markline_opts = opts.MarkLineOpts(
                           data = [opts.MarkLineItem(name = "average",type_ = "average")],
                           label_opts=opts.LabelOpts(is_show=True,position = 'end')
                                                    )
                  )
        .set_series_opts(label_opts=opts.LabelOpts(is_show=False))
        .set_global_opts(
            title_opts=opts.TitleOpts(title="2018-2020项目数增量对比(按周)",
            subtitle=("今天是{}  第{}周").format(dt.now().date(),dt.now().date().isocalendar()[1])),
            datazoom_opts = [
                            opts.DataZoomOpts(is_show=True,
                                        type_ = 'slider',
                                        orient = 'horizontal',
                                        pos_top='29%',
                                        range_start=0,
                                        range_end=100,
                                        xaxis_index=[0,1,2]
                                        ),
                            opts.DataZoomOpts(type_ = 'inside',
                                        orient = 'horizontal',
                                        range_start=0,
                                        range_end=60,
                                        xaxis_index=[0,1,2]
                                             )
                            ],
            tooltip_opts=opts.TooltipOpts(
                                        trigger="axis",
                                        axis_pointer_type="shadow",
                                        background_color="rgba(245, 245, 245, 0.5)",
                                        border_width=1,
                                        border_color="#ccc",
                                        textstyle_opts=opts.TextStyleOpts(color="#000")
                                        ),
            yaxis_opts=opts.AxisOpts(grid_index=0,splitline_opts=opts.SplitLineOpts(is_show=True)),
            xaxis_opts=opts.AxisOpts(
                                    is_show = True,
                                    axistick_opts=opts.AxisTickOpts(is_align_with_label=True)
                                    )
                    )
    )
    bar1 = (
        Bar(init_opts = opts.InitOpts(bg_color='white'))    #0e2147
        .add_xaxis(['week{}'.format(w) for w in range(1,27)])
        .add_yaxis("2018年",
                   [v for v in list(dict_proposed_material_2018.values())],
                   xaxis_index=1,
                   color = '#C23531',
                   gap = 0,
                   markline_opts = opts.MarkLineOpts(
                           data = [opts.MarkLineItem(name = "average",type_ = "average")],
                           label_opts=opts.LabelOpts(is_show=True,position = 'end')
                                                    )
                  )
        .add_yaxis("2019年",
                   [v for v in list(dict_proposed_material_2019.values())],
                   xaxis_index=1,
                   color = '#2F4554',
                   gap = 0,
                   markline_opts = opts.MarkLineOpts(
                           data = [opts.MarkLineItem(name = "average",type_ = "average")],
                           label_opts=opts.LabelOpts(is_show=True,position = 'end')
                                                    )
                  )
        .add_yaxis("2020年",
                   [v for v in list(dict_proposed_material_2020.values())],
                   xaxis_index=1,
                   color = '#61A0A8',
                   gap = 0,
                   markline_opts = opts.MarkLineOpts(
                           data = [opts.MarkLineItem(name = "average",type_ = "average")],
                           label_opts=opts.LabelOpts(is_show=True,position = 'end')
                                                    )
                  )
        .set_series_opts(label_opts=opts.LabelOpts(is_show=False))
        .set_global_opts(
            title_opts=opts.TitleOpts(title="2018-2020提案项目数增量对比(按周)",pos_top="32%",
                                      subtitle=("今天是{}  第{}周").format(dt.now().date(),dt.now().date().isocalendar()[1])),
            tooltip_opts=opts.TooltipOpts(
                                        trigger="axis",
                                        axis_pointer_type="shadow",
                                        background_color="rgba(245, 245, 245, 0.5)",
                                        border_width=1,
                                        border_color="#ccc",
                                        textstyle_opts=opts.TextStyleOpts(color="#000")
                                        ),
            yaxis_opts=opts.AxisOpts(grid_index=0,splitline_opts=opts.SplitLineOpts(is_show=True)),
            xaxis_opts=opts.AxisOpts(
                                    is_show = True,
                                    axistick_opts=opts.AxisTickOpts(is_align_with_label=True)
                                    ),
            legend_opts=opts.LegendOpts(is_show=False)
                    )
    )
    bar2 = (
        Bar(init_opts = opts.InitOpts(bg_color='white'))    #0e2147
        .add_xaxis(['week{}'.format(w) for w in range(1,27)])
        .add_yaxis("2018年",
                   [int(v) for v in list(dict_signed_material_2018.values())],
                   xaxis_index=2,
                   color = '#C23531',
                   gap = 0,
                   markline_opts = opts.MarkLineOpts(
                           data = [opts.MarkLineItem(name = "average",type_ = "average")],
                           label_opts=opts.LabelOpts(is_show=True,position = 'end')
                                                    )
                  )
        .add_yaxis("2019年",
                   [int(v) for v in list(dict_signed_material_2019.values())],
                   xaxis_index=2,
                   color = '#2F4554',
                   gap = 0,
                   markline_opts = opts.MarkLineOpts(
                           data = [opts.MarkLineItem(name = "average",type_ = "average")],
                           label_opts=opts.LabelOpts(is_show=True,position = 'end')
                                                    )
                  )
        .add_yaxis("2020年",
                   [int(v) for v in list(dict_signed_material_2020.values())],
                   xaxis_index=2,
                   color = '#61A0A8',
                   gap = 0,
                   markline_opts = opts.MarkLineOpts(
                           data = [opts.MarkLineItem(name = "average",type_ = "average")],
                           label_opts=opts.LabelOpts(is_show=True,position = 'end')
                                                    )
                  )
        .set_series_opts(label_opts=opts.LabelOpts(is_show=False))
        .set_global_opts(
            title_opts=opts.TitleOpts(title="2018-2020签单物料数增量对比(按周)",pos_top='66%',
            subtitle=("今天是{}  第{}周").format(dt.now().date(),dt.now().date().isocalendar()[1])),
            tooltip_opts=opts.TooltipOpts(
                                        trigger="axis",
                                        axis_pointer_type="shadow",
                                        background_color="rgba(245, 245, 245, 0.5)",
                                        border_width=1,
                                        border_color="#ccc",
                                        textstyle_opts=opts.TextStyleOpts(color="#000")
                                        ),
            yaxis_opts=opts.AxisOpts(grid_index=0,splitline_opts=opts.SplitLineOpts(is_show=True)),
            xaxis_opts=opts.AxisOpts(
                                    is_show = True,
                                    axistick_opts=opts.AxisTickOpts(is_align_with_label=True)
                                    ),
            legend_opts=opts.LegendOpts(is_show=False)
                    )
    )
    grid = (
        Grid(init_opts=opts.InitOpts(height="1300px"))
        .add(bar, grid_opts=opts.GridOpts(pos_bottom="73%"))
        .add(bar1, grid_opts=opts.GridOpts(pos_top='38%',pos_bottom="40%"))
        .add(bar2, grid_opts=opts.GridOpts(pos_top="72%"))
    )
    return grid

#incrementCompareBar().render()

# In[存量分析----可视化] 
def inventoryCompareBar()->Grid:
    bar = (
        Bar(init_opts=opts.InitOpts(bg_color='white'))  
        .add_xaxis([str(i)[5:] for i in get_date_list('2020-01-01',datetime.today().strftime('%Y-%m-%d'))])
        .add_yaxis("进行中",
                   [int(v) for v in ongoing if v is not np.nan],
                   xaxis_index=0,
                   color = '#214761',
                   stack='stack1',
                   category_gap='60%',
                  )
        .add_yaxis("未分配",  
                   [int(v) for v in unallocated if v is not np.nan],
                   xaxis_index=0,
                   color = 'firebrick',
                   stack='stack1',
                   category_gap='60%',
                  )
        .set_series_opts(label_opts=opts.LabelOpts(is_show=False))
        .set_global_opts(
            title_opts=opts.TitleOpts(title="2020项目总数存量(按天)",
                                     subtitle="存量项目=未分配+进行中\n今天是{}  第{}周".format(dt.now().date(),dt.now().date().isocalendar()[1])),
            datazoom_opts = [
                            opts.DataZoomOpts(is_show=True,
                                        type_ = 'slider',
                                        orient = 'horizontal',
                                        pos_top='29%',
                                        range_start=0,
                                        range_end=100,
                                        xaxis_index=[0,1,2]
                                        ),
                            opts.DataZoomOpts(type_ = 'inside',
                                        orient = 'horizontal',
                                        range_start=0,
                                        range_end=60,
                                        xaxis_index=[0,1,2]
                                             )
                            ],
            tooltip_opts=opts.TooltipOpts(
                                        trigger="axis",
                                        axis_pointer_type="shadow",
                                        background_color="rgba(245, 245, 245, 0.5)",
                                        border_width=1,
                                        border_color="#ccc",
                                        textstyle_opts=opts.TextStyleOpts(color="#000")),
            xaxis_opts=opts.AxisOpts(axistick_opts=opts.AxisTickOpts(is_align_with_label=True)),
                    )
    )
    bar1 = (
        Bar(init_opts=opts.InitOpts(bg_color='white'))  
        .add_xaxis([str(i)[5:] for i in get_date_list('2020-01-01',datetime.today().strftime('%Y-%m-%d'))])
        .add_yaxis("进行中",
                   [int(v) for v in proposal_ongoing if v is not np.nan],
                   xaxis_index=1,
                   color = '#214761',
                   stack='stack2',
                   category_gap='60%',
                  )
        .add_yaxis("未分配",

                   [int(v) for v in proposal_unallocated if v is not np.nan],
                   xaxis_index=1,
                   color = 'firebrick',
                   stack='stack2',
                   category_gap='60%',
                  )
        .set_series_opts(label_opts=opts.LabelOpts(is_show=False))
        .set_global_opts(
            title_opts=opts.TitleOpts(title="2020提案项目存量(按天)",pos_top="32%",
                                      subtitle=("今天是{}  第{}周").format(dt.now().date(),dt.now().date().isocalendar()[1])),
            tooltip_opts=opts.TooltipOpts(
                                        trigger="axis",
                                        axis_pointer_type="shadow",
                                        background_color="rgba(245, 245, 245, 0.5)",
                                        border_width=1,
                                        border_color="#ccc",
                                        textstyle_opts=opts.TextStyleOpts(color="#000")
                                        ),
            xaxis_opts=opts.AxisOpts(axistick_opts=opts.AxisTickOpts(is_align_with_label=True))
                    )
    )
    bar2 = (
        Bar(init_opts=opts.InitOpts(bg_color='white'))  
        .add_xaxis([str(i)[5:] for i in get_date_list('2020-01-01',datetime.today().strftime('%Y-%m-%d'))])
        .add_yaxis("进行中",
                   [int(v) for v in signed_ongoing if v == v],
                   xaxis_index=2,
                   color = '#214761',
                   stack='stack3',
                   category_gap='60%',
                  )
        .add_yaxis("未分配",
                   [int(v) for v in signed_unallocated if v == v],
                   xaxis_index=2,
                   color = 'firebrick',
                   stack='stack3',
                   category_gap='60%',
                  )
        .set_series_opts(label_opts=opts.LabelOpts(is_show=False))
        .set_global_opts(
            title_opts=opts.TitleOpts(title="2020签单项目存量(按天)",pos_top="66%",
                                      subtitle=("今天是{}  第{}周").format(dt.now().date(),dt.now().date().isocalendar()[1])),
            tooltip_opts=opts.TooltipOpts(
                                        trigger="axis",
                                        axis_pointer_type="shadow",
                                        background_color="rgba(245, 245, 245, 0.5)",
                                        border_width=1,
                                        border_color="#ccc",
                                        textstyle_opts=opts.TextStyleOpts(color="#000")
                                        ),
            xaxis_opts=opts.AxisOpts(axistick_opts=opts.AxisTickOpts(is_align_with_label=True))
                    )
    )
    grid = (
        Grid(init_opts=opts.InitOpts(height="1300px"))
        .add(bar, grid_opts=opts.GridOpts(pos_bottom="73%"))
        .add(bar1, grid_opts=opts.GridOpts(pos_top='38%',pos_bottom="40%"))
        .add(bar2, grid_opts=opts.GridOpts(pos_top="72%"))
    )
    return grid



# In[总量分析----可视化] 
## 关键时间点折线图
def keyPointLine() -> Line:
    line1= (
        Line()
        .add_xaxis([i for i in Date_List])
        .add_yaxis('2015',
                   [i for i in list(date_2015.values())],
                    xaxis_index=0,
                   color='#C23531',
                   is_connect_nones=True
                  )
        .add_yaxis('2016',
                   [i for i in list(date_2016.values())],
                   xaxis_index=0,
                   color='#2F4554',
                  is_connect_nones=True)
        .add_yaxis('2017',
                   [i for i in list(date_2017.values())],
                   xaxis_index=0,
                   color='#61A0A8',
                  is_connect_nones=True)
        .add_yaxis('2018',
                   [i for i in list(date_2018.values())],
                   xaxis_index=0,
                   color='#008000',
                   is_connect_nones=True)
        .add_yaxis('2019',
                   [i for i in list(date_2019.values())],
                   xaxis_index=0,
                   color='#9400D3',
                   is_connect_nones=True) 
        .add_yaxis('2020',
                   [i for i in list(date_2020.values())],
                   xaxis_index=0,
                   color='#CD853F',
                   is_connect_nones=True) 
        .set_series_opts(label_opts = opts.LabelOpts(is_show = True))
        .set_global_opts( title_opts=opts.TitleOpts(title="项目数总量分析",
                                                    subtitle=("今天是{}  第{}周").format(dt.now().date(),dt.now().date().isocalendar()[1])),
                        xaxis_opts=opts.AxisOpts(is_scale=True),
                        datazoom_opts=[
                                        opts.DataZoomOpts(type_ = 'slider',
                                                        orient = 'horizontal',
                                                        pos_top='29%',
                                                        range_start=15,
                                                        range_end=60,
                                                        xaxis_index=[0,1,2],
                                                        ),
                                        opts.DataZoomOpts(type_ = 'inside',
                                                        orient = 'horizontal',
                                                        range_start=0,
                                                        range_end=60,
                                                        xaxis_index=[0,1,2]
                                                        )])
    )
    line2= (
        Line()
        .add_xaxis([i for i in Date_List])
        .add_yaxis('2015',
                   [i for i in list(Material_2015.values())],
                   xaxis_index=1,
                   color='#C23531',
                   is_connect_nones=True
                  )
        .add_yaxis('2016',
                   [i for i in list(Material_2016.values())],
                   xaxis_index=1,
                   color='#2F4554',
                   is_connect_nones=True)
        .add_yaxis('2017',
                   [i for i in list(Material_2017.values())],
                   xaxis_index=1,
                   color='#61A0A8',
                   is_connect_nones=True)
        .add_yaxis('2018',
                   [i for i in list(Material_2018.values())],
                   xaxis_index=1,
                   color='#008000',
                   is_connect_nones=True)
        .add_yaxis('2019',
                   [i for i in list(Material_2019.values())],
                   xaxis_index=1,
                   color='#9400D3',
                   is_connect_nones=True) 
        .add_yaxis('2020',
                   [i for i in list(Material_2020.values())],
                   xaxis_index=1,
                   color='#CD853F',
                   is_connect_nones=True) 
        .set_series_opts(label_opts = opts.LabelOpts(is_show = True))
        .set_global_opts(title_opts=opts.TitleOpts(title="物料数总量分析",pos_top='32%',
                                                   subtitle=("今天是{}  第{}周").format(dt.now().date(),dt.now().date().isocalendar()[1])),
                        xaxis_opts=opts.AxisOpts(is_scale=True))
    )
    line3= (
        Line()
        .add_xaxis([i for i in Date_List])
        .add_yaxis('2020',
                   [v for v in list(finished_pictures_2019.values())],
                   xaxis_index=2,
                   color='#C23531',
                   is_connect_nones=True) 
        .set_series_opts(label_opts = opts.LabelOpts(is_show = True))
        .set_global_opts(title_opts=opts.TitleOpts(title="完稿画面数总量分析",pos_top='66%',
                                                   subtitle=("今天是{}  第{}周").format(dt.now().date(),dt.now().date().isocalendar()[1])),
                        xaxis_opts=opts.AxisOpts(is_scale=True),
                        legend_opts=opts.LegendOpts(is_show=False))
    )
    grid = (
        Grid(init_opts=opts.InitOpts(height="1300px"))
        .add(line1, grid_opts=opts.GridOpts(pos_bottom="73%"))
        .add(line2, grid_opts=opts.GridOpts(pos_top='38%',pos_bottom="40%"))
        .add(line3, grid_opts=opts.GridOpts(pos_top="72%"))
    )
    return grid

#keyPointLine().render()
# In[AE分析----可视化] 
sort_categories_df_list = deepcopy(categories_df_list)
for i in sort_categories_df_list:
    i['sort_key'] = i['进行中'] + i['已完成'] + i['未完成']
for index, dataframe in enumerate(sort_categories_df_list):
    sort_categories_df_list[index] = dataframe.sort_values(by='sort_key')

def projectDesignedByAE(i:int,categories_df_list):
    bar1 = (  
        Bar(init_opts=opts.InitOpts(width='700px',height='740px'))
        .add_xaxis([n for n in categories_df_list[i].DM])
        .add_yaxis("进行中",[int(j) for j in categories_df_list[i].进行中],
                    xaxis_index=0,
                   stack='stack1',
                   markline_opts=opts.MarkLineOpts(
                                   data=[opts.MarkLineItem(x=12)],
                                   linestyle_opts=opts.LineStyleOpts(type_='dashed',color='orange'),
                                   label_opts=opts.LabelOpts(is_show=True,position='end')))
        .add_yaxis("已完成",[int(j) for j in categories_df_list[i].已完成],
                    xaxis_index=0,                          
                   stack='stack1',
                  )
        .add_yaxis("未完成",[int(j) for j in categories_df_list[i].未完成],
                        xaxis_index=0,
                   stack='stack1',
                  )
        .set_global_opts(title_opts=opts.TitleOpts(title='AE'+categories_list[i]+'数',
                                                   subtitle=("今天是{}  第{}周").format(dt.now().date(),dt.now().date().isocalendar()[1])),             
        xaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(font_size=11,
                                                               interval=0,
                                                               rotate=30,
                                                               vertical_align='middle'),
                            axistick_opts=opts.AxisTickOpts(is_align_with_label=True)),
                         tooltip_opts=opts.TooltipOpts(trigger="axis",
                                                        formatter=utils.JsCode(
                                                               """
                                                              function (params){    
                                                                console.log(params);
                                                                var patten = params[0].axisValue;
                                                                for(var i=0;i<params.length;i++){
                                                                if(params[i].axisIndex === 0)
                                                                {
                                                                params[i].value = params[i].value ;
                                                                };
                                                               
                                                                 patten = patten + '<br>'+ params[i].seriesName + ': ' + params[i].value;
                                                                }
                                                                return patten;
                                                                
                                                                
                                                                }
                                                               """),
                                                      axis_pointer_type="shadow",
                                                      background_color="rgba(245, 245, 245, 0.6)",
                                                      border_width=1,
                                                      border_color="#ccc",
                                                      textstyle_opts=opts.TextStyleOpts(
                                                          color="#000"))
                                                          )
        .set_series_opts(label_opts = opts.LabelOpts(is_show = False,position='right', font_size=15))
           ).reversal_axis()
    return bar1

categories_list=['签单项目','签单物料','完稿画面']
timeline = Timeline(
    init_opts=opts.InitOpts(width='780px',height='770px'))
for index in range(3):
    g = projectDesignedByAE(index,sort_categories_df_list)
    timeline.add(g, time_point=categories_list[index])
    
timeline.add_schema(
    axis_type='category',
    is_auto_play=False,
    is_inverse=True,
    play_interval=2000,
    pos_left="10%",
    pos_right="85%",
#    pos_bottom="75%",
    width=600,
    label_opts=opts.LabelOpts(position='bottom',font_size=12))
#释放空间
del sort_categories_df_list



# In[设计师分析----可视化] 
def designer_analysis()->Bar:
    tl = Timeline(init_opts=opts.InitOpts(width="1200px"))
    for i in range(category_type_Num):
        bar = (
            Bar()
            .add_xaxis([i for i in designers[i]])
            .add_yaxis("未分配", [i for i in distributed[i]],
                       category_gap='40%',
                       stack="stack1")
            .add_yaxis("进行中", [i for i in on_going[i]],
                       category_gap='40%',
                       stack="stack1")
            .add_yaxis("待确稿", [i for i in draft[i]],
                       category_gap='40%',
                       stack="stack1")
            .add_yaxis("已完成", [i for i in completed[i]],
                       category_gap='40%',
                       stack="stack1")
            .add_yaxis("未完成", [i for i in unfinished[i]],
                       category_gap='40%',
                       stack="stack1")
            .extend_axis(
                yaxis=opts.AxisOpts(
                    name="工作量",
                    is_scale = True,
                    type_="value",
                    position="right",
                    offset=20,
                    axisline_opts=opts.AxisLineOpts(
                        linestyle_opts=opts.LineStyleOpts(color="Navy")),
                )
            )
            .set_series_opts(label_opts=opts.LabelOpts(is_show=False,position= 'inside',color = 'white'))
            .set_global_opts(title_opts=opts.TitleOpts(title="设计物料统计",
                                                       subtitle=("今天是{}  第{}周").format(dt.now().date(),dt.now().date().isocalendar()[1])),
                            datazoom_opts = opts.DataZoomOpts(type_ = 'inside',
                                            orient = 'horizontal',range_start=0,range_end=100),
                            tooltip_opts=opts.TooltipOpts(trigger="axis",axis_pointer_type="shadow",
                                            background_color="rgba(245, 245, 245, 0.5)",border_width=1,
                                            border_color="#ccc",textstyle_opts=opts.TextStyleOpts(color="#000")),
                            yaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=True,
                                                     linestyle_opts=opts.LineStyleOpts(
                                                     width=0.5))),
                            xaxis_opts=opts.AxisOpts(
                                axistick_opts=opts.AxisTickOpts(is_align_with_label=True),
                                axislabel_opts=opts.LabelOpts(font_size=11,
                                                            interval=0,
                                                            rotate=25)))
        )
        line = (
            Line()
            .add_xaxis([i for i in designers[i]])
            .add_yaxis('工作量',
                [i for i in workload[i]],
                symbol_size=8,
                yaxis_index=1,
                is_smooth = True,
                itemstyle_opts = opts.ItemStyleOpts(color = "Navy"),
                linestyle_opts=opts.LineStyleOpts(width=2,opacity=0.5,color="Navy"),
                label_opts=opts.LabelOpts(is_show=False))
        )

        bar.overlap(line)
        tl.add_schema(symbol = 'roundRect',pos_top='93%')
        tl.add(bar, "{}".format(category_type[i]))
    return tl

# In[大区分析----可视化] 
##大区分析（五大区项目数饼图和2017-2019大区项目、物料和完稿画面数柱状图）
color_function = """
    function (params) {
       var myColor = ['#E61E24','#f87544','#FF9900','#B6B33F','#259C25'];
       var num = myColor.length;
       return myColor[params.dataIndex % num]
    }
    """
#制作大区分析中的第二个表格
def prop_bar():
    dataframe =  pd.read_excel(PATH+'/2020项目数据.xlsx')[['状态','城市']]
    data = dataframe.merge(city_region_2020, how='left', on='城市')
    data = data.groupby(['大区', '状态'])['项目'].agg('count').unstack(level=-1)
    for i in (set( ['取消','已完成' ,'未完成','进行中','未分配']) - set(data.columns)):
        data[i] = 0
    data['合计行'] = data.sum(axis=1)
#    data.loc['合计列'] = data.sum(axis=0)
    for i in data.columns[:-1]:
        data[i] = data[i]/data['合计行']
    data = data.fillna(0)
    #重新设置索引
    data__ = deepcopy(data)
    data.iloc[0,:] = data__.loc['华中区']
    data.iloc[1,:] = data__.loc['华东区']
    data.iloc[2,:] = data__.loc['华北区']
    data.iloc[3,:] = data__.loc['华南区']
    data.iloc[4,:] = data__.loc['上海区']
    data.index = ['华中区', '华东区', '华北区','华南区', '上海区']
#    data.iloc[:,[0,1,2,3,4]] = (data.iloc[:,[0,1,2,3,4]]*10000/100).applymap(lambda x:'%.2f%%'%x)
    bar = (
            Bar()
            
            .add_xaxis(data.index.to_list())
            .add_yaxis('未分配', data['未分配'].to_list(), stack='stack11',category_gap="30%",xaxis_index=0)
            .add_yaxis('进行中', data['进行中'].to_list(), stack='stack11',category_gap="30%",xaxis_index=0 )
            .add_yaxis('已完成', data['已完成'].to_list(), stack='stack11',category_gap="30%",xaxis_index=0)
            .add_yaxis('未完成', data['未完成'].to_list(), stack='stack11',category_gap="30%",xaxis_index=0)
            .add_yaxis('取消', data['取消'].to_list(), stack='stack11',category_gap="30%",xaxis_index=0)
            .set_series_opts(label_opts=opts.LabelOpts(is_show=False))
            .set_global_opts(xaxis_opts=opts.AxisOpts(max_=1),
                             legend_opts=opts.LegendOpts(pos_top='35%'),
                             title_opts=opts.TitleOpts(title='大区项目进度',
                                                       subtitle=("今天是{}  第{}周").format(dt.now().date(),dt.now().date().isocalendar()[1]),
                                                       pos_top="29%"), 
                                                       )
            ).reversal_axis()
    return bar


def five_regions_grid(i:int) ->Grid:
    bar = (  
        Bar()
        .add_xaxis([n for n in Regions])
        .add_yaxis("2018",[int(j) for j in list(five_regions_df[i]['2018'].values) if j==j],
                   gap=0, color='#C23531',xaxis_index=1)
        .add_yaxis("2019",[int(j) for j in list(five_regions_df[i]['2019'].values) if j==j],
                   gap=0, color='#2F4554',xaxis_index=1)
        .add_yaxis("2020",[int(j) for j in list(five_regions_df[i]['2020'].values) if j==j],
                   gap=0, color='#61A0A8',xaxis_index=1)
        .set_series_opts(label_opts = opts.LabelOpts(is_show =True))
        .set_global_opts(title_opts=opts.TitleOpts(title='2018-2020大区'+three_categories[i]+'数',
                                                       subtitle=("今天是{}  第{}周").format(dt.now().date(),dt.now().date().isocalendar()[1]),
                                                  pos_top='67%'),      
                                xaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(
                                                            font_size=12,
                                                            vertical_align='middle'),
                                axistick_opts=opts.AxisTickOpts(is_align_with_label=True)),
                                legend_opts=opts.LegendOpts(pos_top='68%'),
                             tooltip_opts=opts.TooltipOpts(trigger="axis",
 formatter=utils.JsCode(
                                                               """
                                                              function (params){   
                                                              console.log(params);
                                                                var patten = params[0].axisValue;
                                                                for(var i=0;i<params.length;i++){
                                                                if(params.length == 5){
                                                                params[i].value = Math.floor(params[i].value * 10000) / 100 +'%' ;
                                                               }
                                                               
                                                                 patten = patten + '<br>'+ params[i].seriesName + ': ' + params[i].value;
                                                                }
                                                                return patten;
                                                                }
                                                               """),
                                                          axis_pointer_type="shadow",
                                                          background_color="rgba(245, 245, 245, 0.6)",
                                                          border_width=1,
                                                          border_color="#ccc",
                                                          textstyle_opts=opts.TextStyleOpts(
                                                              color="#000"))
                                )
           )
#    return bar
    pie = (
        Pie(init_opts=opts.InitOpts(bg_color='white'))
        .add(
            "",
            [list(z) for z in zip([i for i in items_region_key],[int(j) for j in items_region_value])],
            radius = [0, '20%'],
            center = ['50%', '18%'],
            itemstyle_opts={
                "normal":{"color":utils.JsCode(color_function)}},
            label_opts=opts.LabelOpts(formatter="{b}\n{c}\n占{d}%"),
        )
        .set_global_opts(title_opts=opts.TitleOpts(title="2020大区项目数分布",
                                                   subtitle="今天是{}  第{}周\n数据截止：{}".format(dt.now().date(),dt.now().date().isocalendar()[1],dt.now().date()),
                                                   title_textstyle_opts=opts.TextStyleOpts(color='#000')),
                        legend_opts=opts.LegendOpts(textstyle_opts=opts.TextStyleOpts(color='#000'),))
    )
    grid = (
        Grid(init_opts=opts.InitOpts(height='1300px'))
#        .add(prop_bar(),grid_opts=opts.GridOpts(pos_top="40%",pos_bottom='35%'))
       .add(bar,grid_opts=opts.GridOpts(pos_top="73%",pos_bottom='6%'))
        .add(prop_bar(),grid_opts=opts.GridOpts(pos_top="40%",pos_bottom='35%'))
         .add(pie,grid_opts=opts.GridOpts(pos_bottom='85%',pos_top='30%'))
#          .add(bar,grid_opts=opts.GridOpts(pos_top="73%",pos_bottom='6%'))
    )

    return grid

three_categories=['项目','物料','完稿画面']
timeline2 = Timeline(init_opts=opts.InitOpts(width='900px',height='1300px'))
for index in range(3):
    b =  five_regions_grid(i=index)
    timeline2.add(b, time_point=three_categories[index])
    
timeline2.add_schema(
    axis_type='category',
    orient="vertical",
    is_auto_play=False,
    is_inverse=True,
    play_interval=2000,
    pos_left="90%",
    pos_top="65%",
    pos_bottom='5%',
    width=70,
    label_opts=opts.LabelOpts(position='right'))


#timeline2.render()


# In[城市分析----可视化] 
## 城市分析（2017-2019项目、物料和完稿画面数TOP10城市;2018-2019提案和签单占比图）
def cities_top10(i:int):
    bar1 = (  
        Bar(init_opts=opts.InitOpts(width="700px", height="800px"))
        .add_xaxis([n for n in three_categories_df[i].city])
        .add_yaxis("2018",[j for j in three_categories_df[i]['2018']],gap=0,color='#C23531')
        .add_yaxis("2019",[j for j in three_categories_df[i]['2019']],gap=0,color='#2F4554')
        .add_yaxis("2020",[j for j in three_categories_df[i]['2020']],gap=0,color='#61A0A8')
        .set_global_opts(title_opts=opts.TitleOpts(title='2018-2020'+three_categories[i]+'数top10城市',
                                                  subtitle="今天是{}  第{}周\n数据截止：{}".format(dt.now().date(),dt.now().date().isocalendar()[1],dt.now().date())),      
        xaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(font_size=12,
                                                               vertical_align='middle'),
                            axistick_opts=opts.AxisTickOpts(is_align_with_label=True)),
                         tooltip_opts=opts.TooltipOpts(trigger="axis",
                                                      axis_pointer_type="shadow",
                                                      background_color="rgba(245, 245, 245, 0.6)",
                                                      border_width=1,
                                                      border_color="#ccc",
                                                      textstyle_opts=opts.TextStyleOpts(
                                                          color="#000")))
        .set_series_opts(label_opts = opts.LabelOpts(is_show =True))
           )
    bar2 = (
        Bar()
        .add_xaxis(cities)
        .add_yaxis("提案2019", 
                   [i for i in proposal_signed_top10['2019_proposal_%']],
                   color='#2F4554',
                   gap=0,
                   stack='stack1')
        .add_yaxis("签单2019", 
                   [i for i in proposal_signed_top10['2019_signed_%']],
                   gap=0,
                    color='#C23531',
                   stack='stack1')
        .add_yaxis("提案2020", 
                   [i for i in proposal_signed_top10['2020_proposal_%']],
                   itemstyle_opts=opts.ItemStyleOpts(color='#C23531'),
                   gap=0,
                   stack='stack2')
        .add_yaxis("签单2020", 
                   [i for i in proposal_signed_top10['2020_signed_%']],
                   itemstyle_opts=opts.ItemStyleOpts(color='#2F4554'),
                   gap=0,
                   stack='stack2')
        .set_series_opts(label_opts = opts.LabelOpts(formatter="{c}%",color="#fff",position="inside"))
        .set_global_opts(title_opts=opts.TitleOpts(title="提案和签单项目占比图",
                                                   subtitle="左边是2018年数据，右边是2019年数据\n今天是{}  第{}周".format(dt.now().date(),dt.now().date().isocalendar()[1]),
                                                   pos_top='52%'),
                        yaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(formatter="{value}%")),
                        tooltip_opts=opts.TooltipOpts(trigger="axis",
                                                      axis_pointer_type="shadow",
                                                      background_color="rgba(245, 245, 245, 0.6)",
                                                      border_width=1,
                                                      border_color="#ccc",
                                                      textstyle_opts=opts.TextStyleOpts(
                                                          color="#000")),
                        legend_opts=opts.LegendOpts(pos_top='52%'))
    )
    grid=(
        Grid(init_opts=opts.InitOpts())
        .add(bar1,grid_opts=opts.GridOpts(pos_bottom="55%"))
        .add(bar2,grid_opts=opts.GridOpts(pos_top="59%"))
    )
    return grid
        

three_categories=['项目','物料','完稿画面']
timeline1 = Timeline(
    init_opts=opts.InitOpts(width="1000px", height="900px"))
for index in range(3):
    g = cities_top10(i=index)
    timeline1.add(g, time_point=three_categories[index])
    
timeline1.add_schema(
    axis_type='category',
    orient="vertical",
    is_auto_play=False,
    is_inverse=True,
    play_interval=2000,
    pos_left="90%",
    pos_bottom="55%",
    pos_top='5%',
    width=70,
    label_opts=opts.LabelOpts(position='right'))

# In[绩效分析----数据处理/可视化] 
##工种明细表
designer_task_df=pd.read_excel(PATH+"/设计师任务明细表.xlsx")
designer_task_df1 = designer_task_df.drop([0],axis=0).fillna('').reset_index(drop=True)
designer_task_df1=designer_task_df1.drop(['工种','工号','区域','岗位'],axis=1)
designer_task_df1['工作量'] = designer_task_df1['工作量'].apply(lambda x: x if x else '')


def designer_task_table() -> Table:
    #设置工作量的显示格式
    def workStyle(list_):
        temp = list(list_)
        for index,value in enumerate(list_):
            if not value:
                temp[index] = '' 
        if list_[-1]:
            temp[-1] = '{:.2f}'.format(list_[-1])
        return temp
        
    table = Table()
    headers = ['姓名','提案物料','创意主视觉', '线下物料',
             '子站设计','H5设计','长图文','插画','三维','前端', 
             '后台','视频','PPT','策划','文案','工作量'] 
    rows = [workStyle(designer_task_df1.iloc[i].values) for i in range(designer_task_df1.shape[0])]
#    rows = [list(designer_task_df1.iloc[i].values) for i in range(designer_task_df1.shape[0])]
    table.add(headers, rows).set_global_opts(
    title_opts=opts.ComponentTitleOpts(title="绩效分析",
                                       subtitle=("今天是{}  第{}周").format(dt.now().date(),dt.now().date().isocalendar()[1]))
    )
    return table

# In[产能分析----数据处理/可视化]
#此段string与原来的生产分析代码一样，只不过是字符串化，采用eval方法进行计算，
#计算得到是生产分析中的第二个表的部分数据
#ps：太懒 所以才用eval的。。。

string = \
'''

def data_liquid() -> Liquid:
    l1 = (
        Liquid()
        .set_global_opts(title_opts=opts.TitleOpts(title="提案",                                
                                                   title_textstyle_opts=opts.TextStyleOpts(font_size=30),
        subtitle="总量：{}\n占比：{}\n同比：{}\n存量：{}\n人均：{}\n满载值：{}\n负载率：{}\n"\
        .format(total_proposal_2019,\
        proposal_ratio_2019,proposal_year_on_year,proposal_stock_2019,\
        proposal_per_capita,3,round(proposal_per_capita/3,2)),
        subtitle_textstyle_opts=opts.TextStyleOpts(font_size=20)))
         )
    l2 = (
        Liquid()
        .set_global_opts(title_opts=opts.TitleOpts(title="签单",
                                                   title_textstyle_opts=opts.TextStyleOpts(font_size=30),
        subtitle="总量：{}\n占比：{}\n同比：{}\n存量：{}\n人均：{}\n满载值：{}\n负载率：{}\n"\
        .format(total_signed_2019,\
        signed_ratio_2019,signed_year_on_year,signed_stock_2019,\
        signed_per_capita,10,round(signed_per_capita/10,2)),
        subtitle_textstyle_opts=opts.TextStyleOpts(font_size=20),pos_left='35%'))
     )
    l3 = (
        Liquid()
        .set_global_opts(title_opts=opts.TitleOpts(title="合计",                                 
                                                   title_textstyle_opts=opts.TextStyleOpts(font_size=30),
        subtitle="总量：{}\n占比：{}\n同比：{}\n存量：{}\n人均：{}\n满载值：{}\n负载率：{}\n"\
        .format(total_proposal_2019+total_signed_2019,round(proposal_ratio_2019+signed_ratio_2019,2),\
        round(proposal_year_on_year+signed_year_on_year,2),round(proposal_stock_2019+signed_stock_2019,2),\
        round(proposal_per_capita+signed_per_capita,2),13,round((proposal_per_capita+signed_per_capita)/13,2),),
        subtitle_textstyle_opts=opts.TextStyleOpts(font_size=20),pos_left='65%'))
    )
    l4 = (
        Liquid()
        .set_global_opts(title_opts=opts.TitleOpts(title="小程序",title_textstyle_opts=opts.TextStyleOpts(font_size=30),                                                  
        subtitle="未分配：{}\n进行中：{}\n已完成：{}\n未完成：{}\n取消：{}\n物料总数：{}"\
        .format(shijie_list[1],shijie_list[2],shijie_list[3],shijie_list[4],shijie_list[5],shijie_list[6]),
        subtitle_textstyle_opts=opts.TextStyleOpts(font_size=20),pos_top='20%'))
    )
    l5 = (
        Liquid()
        .set_global_opts(title_opts=opts.TitleOpts(title="上线快",title_textstyle_opts=opts.TextStyleOpts(font_size=30),
        subtitle="未分配：{}\n进行中：{}\n已完成：{}\n未完成：{}\n取消：{}\n物料总数：{}"\
        .format(shangxiankuai_list[1],shangxiankuai_list[2],shangxiankuai_list[3],\
                shangxiankuai_list[4],shangxiankuai_list[5],shangxiankuai_list[6]),
        subtitle_textstyle_opts=opts.TextStyleOpts(font_size=20),pos_left='35%',pos_top='20%'))
    )
    l6 = (
        Liquid()
        .set_global_opts(title_opts=opts.TitleOpts(title="EVP调研",title_textstyle_opts=opts.TextStyleOpts(font_size=30),
        subtitle="未分配：{}\n进行中：{}\n已完成：{}\n未完成：{}\n取消：{}\n物料总数：{}"\
        .format(EVP[1],EVP[2],EVP[3],EVP[4],EVP[5],EVP[6]),
        subtitle_textstyle_opts=opts.TextStyleOpts(font_size=20),pos_left='65%',pos_top='20%'))
    )
    l7 = (
        Liquid()
        .set_global_opts(title_opts=opts.TitleOpts(title="提案物料",\
        title_textstyle_opts=opts.TextStyleOpts(font_size=30),
        subtitle="总量：{}\n人均：{}\n满载值：{}\n负载率：{}"\
        .format(designer_task_df1_T['总量'][0],designer_task_df1_T_new['人均'][0],\
               designer_task_df1_T_new['满载值'][0],designer_task_df1_T_new['负载率'][0]),
        subtitle_textstyle_opts=opts.TextStyleOpts(font_size=20),pos_top='40%'))
    )
    l8 = (
        Liquid()
        .set_global_opts(title_opts=opts.TitleOpts(title="创意主视觉",\
        title_textstyle_opts=opts.TextStyleOpts(font_size=30),
        subtitle="总量：{}\n人均：{}\n满载值：{}\n负载率：{}"\
        .format(designer_task_df1_T['总量'][1],designer_task_df1_T_new['人均'][1],\
               designer_task_df1_T_new['满载值'][1],designer_task_df1_T_new['负载率'][1]),
        subtitle_textstyle_opts=opts.TextStyleOpts(font_size=20),pos_left='22%',pos_top='40%'))
    )
    l9 = (
        Liquid()
        .set_global_opts(title_opts=opts.TitleOpts(title="签单物料",\
        title_textstyle_opts=opts.TextStyleOpts(font_size=30),
        subtitle="总量：{}\n人均：{}\n满载值：{}\n负载率：{}"\
        .format(designer_task_df1_T['总量'][2],designer_task_df1_T_new['人均'][2],\
               designer_task_df1_T_new['满载值'][2],designer_task_df1_T_new['负载率'][2]),
        subtitle_textstyle_opts=opts.TextStyleOpts(font_size=20),pos_left='45%',pos_top='40%'))
    )
    l10 = (
        Liquid()
        .set_global_opts(title_opts=opts.TitleOpts(title="子站设计",\
        title_textstyle_opts=opts.TextStyleOpts(font_size=30),
        subtitle="总量：{}\n人均：{}\n满载值：{}\n负载率：{}"\
        .format(designer_task_df1_T['总量'][3],designer_task_df1_T_new['人均'][3],\
               designer_task_df1_T_new['满载值'][3],designer_task_df1_T_new['负载率'][3]),
        subtitle_textstyle_opts=opts.TextStyleOpts(font_size=20),pos_left='65%',pos_top='40%'))
    )
    l11 = (
        Liquid()
        .set_global_opts(title_opts=opts.TitleOpts(title="H5设计",\
        title_textstyle_opts=opts.TextStyleOpts(font_size=30),
        subtitle="总量：{}\n人均：{}\n满载值：{}\n负载率：{}"\
        .format(designer_task_df1_T['总量'][4],designer_task_df1_T_new['人均'][4],\
               designer_task_df1_T_new['满载值'][4],designer_task_df1_T_new['负载率'][4]),
        subtitle_textstyle_opts=opts.TextStyleOpts(font_size=20),pos_top='55%'))
    )
    l12 = (
        Liquid()
        .set_global_opts(title_opts=opts.TitleOpts(title="长图文",\
        title_textstyle_opts=opts.TextStyleOpts(font_size=30),
        subtitle="总量：{}\n人均：{}\n满载值：{}\n负载率：{}"\
        .format(designer_task_df1_T['总量'][5],designer_task_df1_T_new['人均'][5],\
               designer_task_df1_T_new['满载值'][5],designer_task_df1_T_new['负载率'][5]),
        subtitle_textstyle_opts=opts.TextStyleOpts(font_size=20),pos_left='22%',pos_top='55%'))
    )
    l13 = (
        Liquid()
        .set_global_opts(title_opts=opts.TitleOpts(title="插画",\
        title_textstyle_opts=opts.TextStyleOpts(font_size=30),
        subtitle="总量：{}\n人均：{}\n满载值：{}\n负载率：{}"\
        .format(designer_task_df1_T['总量'][6],designer_task_df1_T_new['人均'][6],\
               designer_task_df1_T_new['满载值'][6],designer_task_df1_T_new['负载率'][6]),
        subtitle_textstyle_opts=opts.TextStyleOpts(font_size=20),pos_left='45%',pos_top='55%'))
    )
    l14 = (
        Liquid()
        .set_global_opts(title_opts=opts.TitleOpts(title="三维",\
        title_textstyle_opts=opts.TextStyleOpts(font_size=30),
        subtitle="总量：{}\n人均：{}\n满载值：{}\n负载率：{}"\
        .format(designer_task_df1_T['总量'][7],designer_task_df1_T_new['人均'][7],\
               designer_task_df1_T_new['满载值'][7],designer_task_df1_T_new['负载率'][7]),
        subtitle_textstyle_opts=opts.TextStyleOpts(font_size=20),pos_left='65%',pos_top='55%'))
    )
    l15 = (
        Liquid()
        .set_global_opts(title_opts=opts.TitleOpts(title="前端",\
        title_textstyle_opts=opts.TextStyleOpts(font_size=30),
        subtitle="总量：{}\n人均：{}\n满载值：{}\n负载率：{}"\
        .format(designer_task_df1_T['总量'][8],designer_task_df1_T_new['人均'][8],\
               designer_task_df1_T_new['满载值'][8],designer_task_df1_T_new['负载率'][8]),
        subtitle_textstyle_opts=opts.TextStyleOpts(font_size=20),pos_top='70%'))
    )
    l16 = (
        Liquid()
        .set_global_opts(title_opts=opts.TitleOpts(title="后台",\
        title_textstyle_opts=opts.TextStyleOpts(font_size=30),
        subtitle="总量：{}\n人均：{}\n满载值：{}\n负载率：{}"\
        .format(designer_task_df1_T['总量'][9],designer_task_df1_T_new['人均'][9],\
               designer_task_df1_T_new['满载值'][9],designer_task_df1_T_new['负载率'][9]),
        subtitle_textstyle_opts=opts.TextStyleOpts(font_size=20),pos_left='22%',pos_top='70%'))
    )
    l17 = (
        Liquid()
        .set_global_opts(title_opts=opts.TitleOpts(title="视频",\
        title_textstyle_opts=opts.TextStyleOpts(font_size=30),
        subtitle="总量：{}\n人均：{}\n满载值：{}\n负载率：{}"\
        .format(designer_task_df1_T['总量'][10],designer_task_df1_T_new['人均'][10],\
               designer_task_df1_T_new['满载值'][10],designer_task_df1_T_new['负载率'][10]),
        subtitle_textstyle_opts=opts.TextStyleOpts(font_size=20),pos_left='45%',pos_top='70%'))
    )
    l18 = (
        Liquid()
        .set_global_opts(title_opts=opts.TitleOpts(title="PPT",\
        title_textstyle_opts=opts.TextStyleOpts(font_size=30),
        subtitle="总量：{}\n人均：{}\n满载值：{}\n负载率：{}"\
        .format(designer_task_df1_T['总量'][11],designer_task_df1_T_new['人均'][11],\
               designer_task_df1_T_new['满载值'][11],designer_task_df1_T_new['负载率'][11]),
        subtitle_textstyle_opts=opts.TextStyleOpts(font_size=20),pos_left='65%',pos_top='70%'))
    )
    l19 = (
        Liquid()
        .set_global_opts(title_opts=opts.TitleOpts(title="策划",\
        title_textstyle_opts=opts.TextStyleOpts(font_size=30),
        subtitle="总量：{}\n人均：{}\n满载值：{}\n负载率：{}"\
        .format(designer_task_df1_T['总量'][12],designer_task_df1_T_new['人均'][12],\
               designer_task_df1_T_new['满载值'][12],designer_task_df1_T_new['负载率'][12]),
        subtitle_textstyle_opts=opts.TextStyleOpts(font_size=20),pos_top='85%'))
    )
    l20 = (
        Liquid()
        .set_global_opts(title_opts=opts.TitleOpts(title="文案",\
        title_textstyle_opts=opts.TextStyleOpts(font_size=30),
        subtitle="总量：{}\n人均：{}\n满载值：{}\n负载率：{}"\
        .format(designer_task_df1_T['总量'][13],designer_task_df1_T_new['人均'][13],\
               designer_task_df1_T_new['满载值'][13],designer_task_df1_T_new['负载率'][13]),
        subtitle_textstyle_opts=opts.TextStyleOpts(font_size=20),pos_left='22%',pos_top='85%'))
    )
     '''   
def get_first_table_data(string):
    string = string.replace('2019','2020').replace('2018', '2019')
    A = []
    patten =  re.compile('("总量.*?)subtitle_textstyle_opts',re.S)
    for i in re.findall(patten,string):
        list_1 = eval(i.strip(',').strip().replace('\n',' ')[:-1])
        A.append(list_1)
    first_row = ['提案', '签单', '小计']
    first_list = []
    first_list.extend(get_first_table_data_1())
    first_list.append('项目分类 启动 占比 同比 存量 人均 满载值 负载率' .split())
#    second_row = [ '提案物料', '创意主视觉', '签单物料', '子站设计', 'H5设计', '长图文', '插画', '三维', '前端', '后台', '视频', 'PPT', '策划', '文案']
    for index,i in enumerate(A[:3]):
         temp  = re.findall('[\d.]+',i)
         temp[:-1] = [number if float(number) else ''  for number in  temp[:-1]]
         temp[-1]  = ('%.1f%%'  %(100*float(temp[-1]))) if 100*float(temp[-1]) else ''
         temp[1] = (str(temp[1]) + '%') if  str(temp[1]) else ''
         temp[2] = (str(temp[2]) + '%') if  str(temp[2]) else ''
         temp[4] = '%.2f' % float(temp[4])
         if index == 2:
             temp = [temp[0]] + ['']*2 + [temp[3]] + ['']*3
         first_list.append([first_row[index]] + temp)
    first_list.append(['']*8)
    first_list.append('工种分类 总量 占比 同比 存量 人均 满载值 负载率'.split())
    for row in chan_neng_third():
         first_list.append(row)
    return first_list





#得到第一个表格的数据
def get_first_table_data_1():
    wuLiao = pd.read_excel('2020物料数据.xlsx')
    xianMu = pd.read_excel('2020项目数据.xlsx')
    headers = '未分配 进行中  待确稿 已完成 未完成  取消'.split()
    data2 = wuLiao.groupby(['状态'])['数量（个）'].agg('sum')[headers].fillna(0).to_list()
    data1 = xianMu.groupby(['状态'])['项目'].agg('count')[headers].fillna(0).to_list()
    SUM = [xianMu.groupby(['状态'])['项目'].agg('count').sum()
            ,wuLiao.groupby(['状态'])['数量（个）'].agg('sum').sum()]
    contain = pd.DataFrame(np.random.random_sample((2,8)),columns='进度分类 总量 未分配 进行中  待确稿 已完成 未完成  取消'.split())
    contain.iloc[:,0] = '项目数 物料数'.split()
    contain['总量'] = SUM
    contain.iloc[0,2:8] = data1
    contain.iloc[1,2:8] = data2
    contain.iloc[:,1:] = contain.iloc[:,1:].astype(np.float)
    #项目数百分比
    list_1 = [''] + ['%.2f%%' % (100*x) if x else '' for x in (contain.iloc[0].values[1:]/contain.iloc[0,1])]
    #物料数百分比
    list_2 = [''] + ['%.2f%%' % (100*x) if x else '' for x in (contain.iloc[1].values[1:]/contain.iloc[1,1])]
#    contain.loc[:,'分配率'] = (contain['进行中 已完成 未完成  取消'.split()].sum(axis=1)/contain['sum']).apply(lambda x: '%.2f%%' %(x*100))
#    contain.loc[:,'完成率'] = (contain['已完成']/(contain['sum']-contain['取消'])).apply(lambda x:'%.2f%%' %(x*100))
#    for i in contain.index:
#        for j in contain.columns:
#            if not contain.ix[i,j]:
#                contain.ix[i, j] = ''
#    result = [list(i) for i in contain.iloc[:,:-1].values]
#    result.append(['']*8)
    result = []
    result.append(contain.iloc[0].to_list())
    result.append(list_1)
    result.append(contain.iloc[1].to_list())
    result.append(list_2)
    result.append(['']*8)
    fun(result)
    result = [[one if one else '' for one in list_item] for list_item in result]
    return result


#产能分析第三个表
def chan_neng_third():
    cols1 = ['状态'] + '主视觉设计师	线下设计师	线上设计师	插画	子站	文案	视频	PPT	前端工程师'.split() + ['数量（个）']
    data_wuliao = pd.read_excel(PATH+'/2020物料数据.xlsx',usecols=cols1)
    titles = ['提案', '平面', '网页', '其他']
    result = dict.fromkeys(titles, 0)
    for title in titles:
        names = df_designer[df_designer['类别'] == title]['姓名'].to_list()
        sum_1 = 0
        for job in '主视觉设计师 线下设计师 线上设计师	 插画	子站	 文案	视频	PPT	前端工程师'.split():
                data = data_wuliao[[job,'数量（个）']]
                data = data[data[job].isin(names)]
                if data[job].any():
                    sum_1 += data['数量（个）'].sum()
                else:
                    sum_1 += 0
        result[title] = sum_1
        #总量  ['提案', '平面', '网页', '其他']
    result_1 = [int(i) for i in result.values()]
    #以下计算存量
    data_wuliao = data_wuliao[data_wuliao['状态'] == '进行中']
    result = dict.fromkeys(titles, 0)
    for title in titles:
        names = df_designer[df_designer['类别'] == title]['姓名'].to_list()
        sum_1 = 0
        for job in '主视觉设计师 线下设计师 线上设计师	插画	子站	文案	视频	PPT	前端工程师'.split():
                data = data_wuliao[[job,'数量（个）']]
                data = data[data[job].isin(names)]
                if data[job].any():
                    sum_1 += data['数量（个）'].sum()
                else:
                    sum_1 += 0
        result[title] = sum_1
    #存量  ['提案', '平面', '网页', '其他']
    result_2 = [int(i) for i in result.values()]
    #人数为 ['提案', '平面', '网页', '其他']
    numbers = df_designer.groupby('类别')['姓名'].count()[titles].to_list()
    #满载值
    result_3 = [3,3,2,2]
    #人均
    mean = ['']*4
    #负载率
    pecent = ['']*4
    #占比
    rati = [''] * len(titles)
    #同比
    same_rati = [''] * len(titles)
    for index in range(len(titles)):
        mean[index] = result_2[index]/numbers[index]
        pecent[index] = mean[index]/result_3[index]
    mean = ['%.2f' %(x) for x in mean]
    pecent = ['%.2f%%' %(x*100) for x in pecent]
    result = [['提案'],['平面'],['网页'],['其他']]
    for row in range(len(titles)):
        result[row].append(result_1[row])
        result[row].append(rati[row])
        result[row].append(same_rati[row])            
        result[row].append(result_2 [row])
        result[row].append(mean[row])            
        result[row].append(result_3[row]) 
        result[row].append(pecent[row])
    return result

#def get_first_table_data(string):
#    
#    A = []
#    patten =  re.compile('("总量.*?)subtitle_textstyle_opts',re.S)
#    for i in re.findall(patten,string):
#        list_1 = eval(i.strip(',').strip().replace('\n',' ')[:-1])
#        A.append(list_1)
#    first_row = ['提案', '签单', '小计']
#    first_list = []
#    first_list.extend(get_first_table_data_1())
#    first_list.append('项目分类 总量 占比 同比 存量 人均 满载值 负载率' .split())
##    second_row = [ '提案物料', '创意主视觉', '签单物料', '子站设计', 'H5设计', '长图文', '插画', '三维', '前端', '后台', '视频', 'PPT', '策划', '文案']
#    for index,i in enumerate(A[:3]):
#         temp  = re.findall('[\d.]+',i)
#         temp[:-1] = [number if float(number) else ''  for number in  temp[:-1]]
#         temp[-1]  = ('%.1f%%'  %(100*float(temp[-1]))) if 100*float(temp[-1]) else ''
#         temp[1] = (str(temp[1]) + '%') if  str(temp[1]) else ''
#         temp[2] = (str(temp[2]) + '%') if  str(temp[2]) else ''
#         temp[4] = '%.2f' % float(temp[4])
#         if index == 2:
#             temp = [temp[0]] + ['']*2 + [temp[3]] + ['']*3
#         first_list.append([first_row[index]] + temp)
#    first_list.append(['']*8)
#    first_list.append('工种分类 总量 占比 同比 存量 人均 满载值 负载率'.split())
#    for row in chan_neng_third():
#         first_list.append(row)
#    return first_list



def make_first_table()->Table:  
    table = Table()
    headers = '进度分类 总量 未分配 进行中 待确稿 已完成 未完成  取消' .split()
    rows = get_first_table_data(string)
    table.add(headers, rows).set_global_opts(
    title_opts=opts.ComponentTitleOpts(title="产能分析",
                                       subtitle=("今天是{}  第{}周").format(dt.now().date(),dt.now().date().isocalendar()[1]))
    )
    return table


# In[物料分析----可视化]
#物料分析
def product_analysis_table() -> Table:
#    product_analysis_df_sorted.未分配 = product_analysis_df_sorted[['排队中','未分配']].replace('',0).sum(axis=1).replace(0, '')
#    product_analysis_df_sorted = product_analysis_df_sorted.drop(columns=['排队中'])
    table = Table()
    headers = list(product_analysis_df_sorted.columns)
    rows = [list(product_analysis_df_sorted.iloc[i].values)\
            for i in range(product_analysis_df_sorted.shape[0])]
    rows = [HE_JI] + rows
    fun(rows)
    table.add(headers, rows).set_global_opts(
    title_opts=opts.ComponentTitleOpts(title="2020物料分析（个）",
                                       subtitle=("今天是{}  第{}周").format(dt.now().date(),dt.now().date().isocalendar()[1]))
    )
    return table



# In[种类分析----数据处理/可视化]
def species_analysis():
#    进行拼音转换
    def hanzi2pinyin(columns):
#        return (i[-1] for i in sorted((lazy_pinyin(i),i) for i in columns))
        columns = [i.replace('曾', '站-') for i in columns]
        return (i.replace('站-', '曾') for  i in sorted(columns, key=lazy_pinyin))
   
    def get_AE_data():
        df = pd.read_excel('2020物料数据.xlsx')
        df3 = df[['DM', '小类','数量（个）']]
        df3 = df3[df3['小类'].isin(['排版类', '文字类', '网页&子站', '设计类'])]
        items_dm = df3.groupby(['DM','小类'])['数量（个）'].count()
        data = items_dm.unstack(level=-1).fillna(0)
        data['sum'] = np.sum(data,axis=1)
        pinyin_sorted = list(hanzi2pinyin(data.index))
        data = data.loc[pinyin_sorted,:]
        for i in data.columns:
            data[i]  = data[i]/data['sum']
        data = data.T
        return data

    def bar_stack0(dataframe) -> Bar:
        w = dataframe
        c = (
            Bar(init_opts=opts.InitOpts(width='780px',height='780px',theme=ThemeType.WHITE))
            .add_xaxis(w.columns.to_list())
            .add_yaxis( w.index[2],dataframe.iloc[2,:].to_list(),xaxis_index=0, stack="stack0",itemstyle_opts=opts.ItemStyleOpts(color='rgb(221,182,53)'))
            .add_yaxis( w.index[1],dataframe.iloc[1,:].to_list(), xaxis_index=0,stack="stack0",itemstyle_opts=opts.ItemStyleOpts(color='rgb(89,74,131)'))
            .add_yaxis( w.index[0], dataframe.iloc[0,:].to_list(), xaxis_index=0,stack="stack0",itemstyle_opts=opts.ItemStyleOpts(color='rgb(162,207,90)'))
            .add_yaxis( w.index[3],dataframe.iloc[3,:].to_list(),xaxis_index=0, stack="stack0",itemstyle_opts=opts.ItemStyleOpts(color='rgb(75,139,174)'))
            .set_global_opts(title_opts=opts.TitleOpts(title='百分比',
                                                       subtitle=("今天是{}  第{}周").format(dt.now().date(),dt.now().date().isocalendar()[1])),                        
                            datazoom_opts = [opts.DataZoomOpts(type_='slider',
                                                              is_show=True,  
                                                              orient = 'vertical',
                                                              yaxis_index=0,
                                                             
                                                          ),
                                                opts.DataZoomOpts(type_='inside',
                                                              is_show=True,
                                                             yaxis_index=0,
                                                              orient = 'vertical',
                                                            )
                                                ], 
                             tooltip_opts=opts.TooltipOpts(trigger="axis",
                                                           formatter=utils.JsCode(
                                                               """
                                                              function (params){    
                                                                console.log(params);
                                                                var patten = params[0].axisValue;
                                                                for(var i=0;i<params.length;i++){
                                                                if(params[i].axisIndex === 0)
                                                                {
                                                                params[i].value = Math.floor(params[i].value * 10000) / 100 +'%' ;
                                                                };
                                                               
                                                                 patten = patten + '<br>'+ params[i].seriesName + ': ' + params[i].value;
                                                                }
                                                                return patten;
                                                                
                                                                
                                                                }
                                                               """),
                                                          axis_pointer_type="shadow",
                                                          background_color="rgba(245, 245, 245, 0.6)",
                                                          border_width=1,
                                                          border_color="#ccc",
                                                          textstyle_opts=opts.TextStyleOpts(
                                                              color="#000"))
                                                              , legend_opts=opts.LegendOpts())
            .set_series_opts(label_opts = opts.LabelOpts(is_show = False))
               ).reversal_axis()
    
        return c
    data = get_AE_data()
    a = bar_stack0(data)
    return a


# In[产品分析----数据处理/可视化]
#产品分析板块

#得到每一种种类的分割处的index
def get_split_index(dataframe):
    temp = dataframe['种类'].to_list()
    split_index = []
    for i in range(1,len(temp)):
        if temp[i-1] != temp[i]:
            split_index.append(i)
        elif i == (len(temp)-1):
            split_index.append(i+1)
    return split_index

def get_table_data():
    data = pd.read_excel('2020物料名称匹配表.xlsx')
    #总表便于统计数量
    data_stata = pd.read_excel('2020物料数据.xlsx')
    ming_xi= data['明细.1'].dropna()
    kinds = kinds = data['种类'][data['明细.1'].notna()].ffill()
    #物料名称 -> 明细
    map_realation_first = dict(zip(data['物料名称'],data['明细']))
    #明细 -》 种类
    map_realation_second = dict(zip(ming_xi ,kinds))
    data_stata['下单月份'] = data_stata['下单日期'].apply(lambda x:dt.strptime(x, '%Y-%m-%d')).dt.month
    data_merge = data_stata[['物料名称','下单月份','数量（个）']]
    data_merge['明细']  = data_merge['物料名称'].apply(lambda x:map_realation_first.get(x.strip(), '无'))
    data_merge = data_merge[data_merge['明细']!='无']
    temp = data_merge.groupby(by=['明细','下单月份'])['数量（个）'].apply(np.sum)
    temp =  temp.unstack(level=-1)
    for i in range(1,13):
        try:
            temp[i]
        except KeyError:
            temp[i] = np.nan
            
    target1 = temp.drop(axis=1,labels=[7,8,9,10,11,12]).fillna('')
#    因为提供的2019物料名称匹配表中可能有平面宝设计的明细，将其删除利用单独的公式进行计算
    try:
        target1.drop(index='平面宝设计',inplace=True)
    except:
        pass
    try:
        target1[12].fillna('')
    except KeyError:
        target1[12] = ''
    #单独处理 当月的对折页+三折页+四折页
    params = ['对折页','三折页','四折页']
    sum_ = data_stata[data_stata['物料名称'].isin(params)]
    #得到一个映射 map_month：键为月份  值为：当月的对折页+三折页+四折页的值
    map_month =dict()
    wanted = sum_.groupby('下单月份')['物料名称'].agg('count')
    #选择1-6月的数据
    for i in range(1,7):
        map_month[i] =  wanted.get(i,0)
        
    target1 = pd.concat([target1, pd.DataFrame([(map_month.values())],index=['平面宝设计'],columns=map_month.keys())])
    target = target1.reset_index()
    target.rename(columns={'index':'明细'},inplace=True)

    #汇总表格
    total_table = data[['明细.1']].rename(columns={'明细.1':'明细'}).dropna()
#    total_table = total_table.reset_index(drop=True)
    for i in range(total_table.shape[0]):
        total_table.loc[i,'种类'] = map_realation_second.get(total_table.ix[i,'明细'],np.nan)
    total_table = total_table.merge(target,on='明细',how='left').fillna('')
    #    计算   平面宝数量=（当月的对折页+三折页+四折页）-电脑端网站-响应式网站
    temp_1 = total_table.set_index('明细')
    #电脑端网站计数
    count_1 = data_merge[data_merge['物料名称']=='子站（定制）'].groupby('下单月份')['物料名称'].count()
    #响应式网站计数
    count_2 = data_merge[data_merge['物料名称']=='子站适配移动端'].groupby('下单月份')['物料名称'].count()
    for i in range(1,7):
        exc_zero = temp_1.loc['平面宝设计',i] -  count_1.get(i,0) - count_2.get(i,0)
        temp_1.loc['平面宝设计',i] = exc_zero if exc_zero  else ''
#    for i in range(1,7):
#        temp_1.loc['平面宝设计',i] =  ''
    total_table = temp_1.reset_index()
    total_table['合计'] = total_table.replace('',0).sum(axis=1)
    total_table = total_table.replace(0,'')
    #选择1月-6月的数据
    headers = ['种类','明细', 1,2, 3,4, 5, 6,'合计']
    total_table = total_table[headers]
    return total_table.values,get_split_index(total_table)



def designer_product_analysis():    
    HEADER = ['种类','明细','1月','2月','3月','4月', '5月', '6月','合计']
    rows,split_index = get_table_data()
    fun(rows)
    table = Table()
    table.add(HEADER, rows).set_global_opts(
        title_opts=ComponentTitleOpts(title="产品分析（个）", subtitle=("今天是{}  第{}周").format(dt.now().date(),dt.now().date().isocalendar()[1]))
    )
    return table,split_index


table_product_analysis, split_index= designer_product_analysis()

################################################################################

############################################
# In[图片进行base64编码]
def encode_image(filename):
    ext = filename.split(".")[-1]
    with open(filename, "rb") as f:
        img = f.read()
    data = base64.b64encode(img).decode()
    src = "data:image/{ext};base64,{data}".format(ext=ext, data=data)
    return src

############################################
# In[视觉分析----数据处理/可视化]
#视觉分析，插入图片
#为设计分析插入图片，
#注意将所要插入的图片存入与PATH目录下的子文件夹（images）中
    
#拼接出符合要求的的html
def prepare_imgs():

    #取出所有的图片
    image_names = [('images/' + name) for name in os.listdir('images') for item in ['.jpg', '.JPG'] if
                       os.path.splitext(name)[1] == item]
    #图片的base64编码列表
    jpgs_base = [encode_image(jpg_name) for jpg_name in image_names]

    pattern = \
    """
        <div class="grid-sizer">
            <div class="grid-item">
                <img class="pimg" src="{}" />                
            </div>
        </div>
        """.format
    
    result = '<div class="grid" id="container">'  + \
            ''.join(pattern(i) for i in jpgs_base) + \
            """
            </div>
    <!-- 用于显示放大的突破 -->
    <div id="outerdiv" style="position:fixed;top:0;left:0;background:rgba(0,0,0,0.7);z-index:2;width:100%;height:100%;display:none;"><div id="innerdiv" style="position:absolute;"><img id="bigimg" style="border:5px solid #fff;" src="" /></div></div>  
    
     """
    return result


#创建视觉分析的tab选项卡
def image_design() -> echart_image:
    image = echart_image()

    img_src = (
        ''
    )
    image.add(
        src=img_src,
        style_opts={"width": "1400px", "height": "5500px", "style": "margin-top: 20px"},
    ).set_global_opts(
        title_opts=ComponentTitleOpts(title="视觉分析", 
                                      subtitle=("今天是{}  第{}周").format(dt.now().date(),dt.now().date().isocalendar()[1]))
    )
    return image


################################################################################
# In[反馈分析----数据处理/可视化]
#保存的二维码图片位置---IMAGE_SAVE_PATH_for_QR
IMAGE_SAVE_PATH_for_QR  = os.path.join(PATH, '二维码.jpg')

def image_QR(IMAGE_SAVE_PATH) -> echart_image:
    image = echart_image()

    img_src = (
        encode_image(IMAGE_SAVE_PATH_for_QR) #图片进行base64编码
    )
    image.add(
        src=img_src,
        style_opts={"width": "450px", "height": "670px", "style": "margin-top: 20px"},
    ).set_global_opts(
        title_opts=ComponentTitleOpts(title="反馈分析", 
                                      subtitle=("今天是{}  第{}周").format(dt.now().date(),dt.now().date().isocalendar()[1]))
    )
    return image



# In[函数---为所有图片设置居中]
#为所有图片设置居中
def QR_mediate():
    with open('designdata.html', 'r', encoding='utf8') as f:
        text = f.read()
    #控制正则匹配的文本区间
    index = 0
    a = []
    patten =  re.compile('<img(.*?)/>', re.S)
    while True:
        one = patten.search(text[index:])
        try:
            img_kept = one.group(0)
        except:
            break
        string = """<div class="demo">
            %s
            </div>""" % img_kept
        a.append((img_kept, string))
        index = text.find(img_kept) +100     
    for i,j in a:
       text=  text.replace(i,j,1)
    with open('designdata.html', 'w', encoding='utf8') as f:
        f.write(text)

########################################################

# In[table选项卡网页是否展现设置处]
# 把所有图表添加到选项卡中
tab = Tab(page_title = '设计中心数据分析')
tab.add(image_design(), '视觉分析')
tab.add(make_first_table(),"产能分析")
tab.add(table_product_analysis,'产品分析')
tab.add(keyPointLine(),"总量分析")
tab.add(incrementCompareBar(),"增量分析")
tab.add(inventoryCompareBar(),"存量分析")
tab.add(species_analysis(),"种类分析")
tab.add(product_analysis_table(),"物料分析")
tab.add(timeline, "AE分析")
tab.add(designer_analysis(), "设计师分析")
tab.add( designer_task_table(), "绩效分析")
tab.add(timeline2,'大区分析')
tab.add(timeline1,"城市分析")
tab.add(image_QR(IMAGE_SAVE_PATH_for_QR), '反馈分析')
tab.render_notebook()
tab.render(PATH+'/designdata.html')



# =============================================================================
#**********************   对网页html进行操作    *******************************
# =============================================================================
# In[利用CSS进行布局]   
#写入首行冻结的样式
css_first_row = """

     *{
            box-sizing:border-box;
          }
      #container{
           position: relative;
           margin : auto;
           width:80%;
        }
        .grid-item{
            width:50%;
            height:auto;
            border: 1px solid #cccccc;
            box-shadow: 0 0 5px #cccccc;
        }
        .grid-sizer{
            float: left;
        }
        .grid-item img{
            width:100%;
            height: auto;
        }
     #head-table_产品分析（个）{
                left: 35px         
            }
       #head-table_产品分析（个） th{
       text-align:center
       }
     #head-table_绩效分析{
                left: 35px          
            }
       #head-table_绩效分析 th{
       text-align:center
       }
     #head-table_物料分析（个）{
                left: 35px          
            }
       #head-table_物料分析（个） th{
       text-align:center
       }
"""

css_style = """
         .demo{
                vertical-align: middle;
                text-align: center;
            }
        table#产能分析 tr td:first-child{
         text-align:left
        }
        table#产能分析 tr td{
         text-align:right
        }
        table#产能分析 tr:nth-child(1) th{
         text-align:center
        }
        table#产能分析 tr:nth-child(7) td{
         text-align:center
        }
        table#产能分析 tr:nth-child(7) td{
         font-weight:bold
        }
        table#产能分析 tr:nth-child(12) td{
         text-align:center
        }
        table#产能分析 tr:nth-child(12) td{
         font-weight:bold
        }
        table#产品分析（个） tr th{
         text-align:center
        }
        table#物料分析（个） tr th{
        text-align:center
        }
        table#物料分析（个） tr td{
         text-align:right
        }
        table#物料分析（个） tr td:first-child{
         text-align:left
        }
        table#绩效分析 tr td{
         text-align:right
        }
        table#绩效分析 tr td:first-child{
         text-align:left
        }
        %s
        """ % css_first_row

with open('designdata.html','r+',encoding='utf8') as f:
    text = f.read()
    text = text.replace('<style>','<style>'+css_style,1)
with open('designdata.html','w',encoding='utf8') as f:    
    f.write(text)

# In[JS相关内容] 
#JS注入相关操作
################################################################################
#更改单元格的JS代码
js_func_1= \
"""
String.prototype.format = function () {
    var values = arguments;
    return this.replace(/\{(\d+)\}/g, function (match, index) {
        if (values.length > index) {
            return values[index];
        } else {
            return "";
        }
    });
};
var length = document.querySelectorAll('table#产能分析 tr').length
for(var i=13;i<=length;i++){
 for(var j=3;j<5;j++){
try{
var c = document.querySelector('table#产能分析 tr:nth-child({0}) td:nth-child({1})'.format(i,j)).bgColor = "#DCDCDC";
}catch(TypeError)
{;};
}
};
for(var i=5;i<=length;i++){
try{
var c = document.querySelector('table#产能分析 tr:nth-child({0}) td:nth-child({1})'.format(i,8));
var s = c.innerHTML;
if(!s){;}
else if(parseFloat(s.substr(0,s.length-1)/100)>1.2){
    c.bgColor="#FF0000";
}else if(parseFloat(s.substr(0,s.length-1)/100)>0.8)
{c.bgColor='#33CCFF';}
}catch(TypeError)
{;};
}



"""
#####################################################################
#合并单元格的js代码
js_func_2 = \
"""
var array = %s;
/**
 * 合并单元格(如果结束行传0代表合并所有行)
 * @param table1    表格的ID
 * @param startRow  起始行
 * @param endRow    结束行
 * @param col   合并的列号，对第几列进行合并(从0开始)。第一行从0开始
 */
function mergeCell(table1, startRow, endRow, col) {
    var tb = document.getElementById(table1);
    if(!tb || !tb.rows || tb.rows.length <= 0) {
        return;
    }
    if(col >= tb.rows[0].cells.length || (startRow >= endRow && endRow != 0)) {
        return;
    }
    if(endRow == 0) {
        endRow = tb.rows.length - 1;
    }
    for(var i = startRow; i < endRow; i++) {
        if(tb.rows[startRow].cells[col].innerHTML == tb.rows[i + 1].cells[col].innerHTML) { //如果相等就合并单元格,合并之后跳过下一行
            tb.rows[i + 1].removeChild(tb.rows[i + 1].cells[col]);
            tb.rows[startRow].cells[col].rowSpan = (tb.rows[startRow].cells[col].rowSpan) + 1;
        } else {
            mergeCell(table1, i + 1, endRow, col);
            break;
        }
    }
}
"""%([0] + split_index)

def prepare_js_fun_2(id_):
        global js_func_2
        str1 = """
        for(var i=0;i<array.length-1;i++){
        mergeCell('%s',array[i],array[i+1],0);
        }
        """% id_
        js_func_2 = js_func_2 + str1
        
#####################################################################
#点击字段进行降序排序
#count表示用一次js_func_3就自加一次防止函数重名
#注意 ： 第一个表示函数的调用次数， 第四个格式字符串为唯一的title即可个：绩效分析或者物料分析
COUNT = 1
js_func_3 = \
"""
    function SortTable%s(obj){
       if (typeof back_.pre != 'string')
    {
    back_['pre'].bgColor = 'white';
    };
    obj.bgColor = '#FFFF00';

        %s
        %s
        %s
        var tds = document.getElementsByName("%s" +  /\d+/.exec(obj.id)[0]);
        //得到当前传入对象的那一列
        var columnArray=[];
        for(var i=0;i<tds.length;i++){
            columnArray.push(tds[i].innerHTML);
        }//当前那一列都写入column这个栈，是逆序的
        var orginArray=[];
        for(var i=0;i<columnArray.length;i++){
            orginArray.push(columnArray[i]);
        }//将这一列的内容再存储一遍，一会原来列表修改以后，
        //通过比对值的方式对应到当前行的内容，实现同行内容一起修改
        columnArray.sort(fun);   //排序后的新值，只排序了当前列
        for(var i=0;i<columnArray.length;i++){
            for(var j=0;j<orginArray.length;j++){
                if(orginArray[j]==columnArray[i]){
                    %s
                    orginArray[j]=null;
                    break;
                }
            }
        }
        back_.pre = obj;
    }
"""

############################################
#为一些字段添加注释详细信息
js_func_4 = \
"""
document.querySelector('table#产能分析 tr:nth-child(1) th:nth-child(3)').title="目前项目数占去年项目总数的比例";
document.querySelector('table#产能分析 tr:nth-child(1) th:nth-child(4)').title="目前项目数同比去年时间点";

"""

############################################
#为产品分析单元格设置对齐格式
js_func_5 = \
"""
(function(){
var str = document.querySelector('table#产品分析（个）').innerHTML;
zz = [];
sp = str[0]
list = str.split(sp);
for(var i=0;i<list.length;i++){
 if(RegExp('<td>([\\\\d\\\\.%]{1,})','g').test(list[i])){
    list[i] = list[i].replace(RegExp('<td>([\\\\\\d\\\\\\.%]{1,})','g'),'<td align="right">'+RegExp["$1"])}
zz.push(list[i])
}
document.querySelector('table#产品分析（个）').innerHTML = zz.join(sp);
})()
"""


############################################
#为首行冻结功能相关的js与jquery

#写入head中的Jquery链接
string_head =\
"""
        <link rel="stylesheet" href="https://cdn.bootcss.com/bootstrap/3.3.7/css/bootstrap.min.css" integrity="sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u" crossorigin="anonymous">
        <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
        <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
        <!--[if lt IE 9]>
            <script src="https://cdn.bootcss.com/html5shiv/3.7.3/html5shiv.min.js"></script>
            <script src="https://cdn.bootcss.com/respond.js/1.4.2/respond.min.js"></script>
        <![endif]-->
"""

#利用jquery首行固定js
#封装成shou_hen( thead_, tbody_)
#param： thead_ 为浮动首行的table id
#         tbody_ 为表格主题的table id

#<script src="https://cdn.bootcss.com/jquery/1.12.4/jquery.min.js"></script>
js_func_6 = \
"""
       <!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
        <!-- Include all compiled plugins (below), or include individual files as needed -->
        <script src="https://cdn.bootcss.com/bootstrap/3.3.7/js/bootstrap.min.js" integrity="sha384-Tc5IQib027qvyjSMfHjOMaLkfuWVxZxUPnCJA7l2mCWNIpG9mGCD8wGNIcPD7Txa" crossorigin="anonymous"></script>
        
         <!-- 封装成shou_hen( thead_, tbody_)
         param： thead_ 为浮动首行的table id
        tbody_ 为表格主题的table id -->
        
        <script type="text/javascript">
        function shou_hen( thead_, tbody_){
           eval(`$(function(){
                $('${thead_}').find('th').each(function(){
                    $(this).width($('${tbody_}').find('th:eq(' + $(this).index() + ')').width()+2);
                });   
 $('${thead_}').css('background-color', 'white');          
                $(window).scroll(function (event) {
                    var scroll = $(window).scrollTop();
                    if (scroll > 50) {
                        $('${thead_}').css('display', 'block');
                        for(var i=0;i<$('${tbody_}').find('th').length;i++){
                        if ($('${tbody_}').find('th')[i].bgColor && $('${tbody_}').find('th')[i].bgColor != 'white')
                            $('${thead_}').find('th')[i].bgColor = $('${tbody_}').find('th')[i].bgColor;
                        }
                    } else {
                        $('${thead_}').css('display', 'none');
                        for(var i=0;i<$('${thead_}').find('th').length;i++){
                        if ($('${thead_}').find('th')[i].bgColor && $('${thead_}').find('th')[i].bgColor != 'white')
                            $('${thead_}').find('th')[i].bgColor = 'white';
                        }
                    }
                });
            });`)
}
</script>
"""


############################################
# In[函数-----为板块加上一个唯一的ID]
#给每一个板块加上一个唯一的id便于操控js代码
#列如
#给产能分析添加id确保改变的表格是唯一的
#id 加在tagname=table的节点上
def make_up(filename,name=None,js_func=None):
    with open('{}'.format(filename),'r',encoding='utf8') as f:
        w = f.read()
    if  name is not None:
        patten = re.compile('> {}</p>.*?<table'.format(name),re.S)
        text = w = re.sub(patten,re.search(patten,w).group(0) + ' id="{}"'.format(re.search('[^\d]+',name).group()),w) 
    if js_func is not None:
        text = w.replace('</html>','\n<script>{}\n</script></html>'.format(js_func))
    with open('{}'.format(filename),'w',encoding='utf8') as f:
        f.write(text)

############################################
# In[table点击表头进行排序]
#点击字段名称，按大到小排序
def prepare_html(title,filename):
    with open('{}'.format(filename),'r',encoding='utf8') as f:
        str2 = f.read()
    str3 = re.search('{}</p>.*?<table.*?</table>'.format(title),str2,re.S).group(0)
    patten = '(<th>?(.*?)</th>)'
    list_1 = re.findall(patten,str3,re.S)
    string =str3
    for index, (pre, value) in enumerate(list_1):
        pa = '<th id="%s" onclick="SortTable%s(this)" class="as">%s</th>' %((title+str(index)), COUNT,value)
        string = string.replace(pre,pa,1)
    patten = '<tr>(.*?)</tr>'
    list_2 = re.findall(patten, string, re.S)[1:]
    for str_1 in list_2:
       temp = str_1
       for index, (pre, value) in enumerate(re.findall('(<td>(.*?)</td>)',str_1, re.S)):
              pa = '<td name="%s">%s</td>' %((title+str(index)), value)
              temp = temp.replace(pre, pa,1)
       string = string.replace(str_1, temp,1)
    length = index + 1
    text = str2.replace(str3, string,1)
    with open('{}'.format(filename),'w',encoding='utf8') as f:
        f.write(text)
    return length

#补充完善js_func_3
def finish_js_func_3(title,length):
    global COUNT
    format_tuple = (
            COUNT,
        '\n'.join('var td%ds=document.getElementsByName("%s");' %(i,title+str(i)) for i in range(length)),
         '\n'.join('var tdArray%s=[];' %i for i in range(length)),
         '\n'.join("""for(var i=0;i<td%ds.length;i++){
            tdArray%s.push(td%ds[i].innerHTML);
        }"""%(i,i,i) for i in range(length)),
            title,
            '\n'.join('document.getElementsByName("%s")[i].innerHTML=tdArray%s[j];'%(title+str(i),i) for i in range(length))
            
)
    str_1 = """var back_ = {pre:''};
    function fun(x1,x2){
    if ( x1.substring(x1.length-1) == "%"){
      x1 = x1.substring(0,x1.length-1);
    }
    if ( x2.substring(x2.length-1) == "%"){
      x2 = x2.substring(0,x2.length-1);
    }
    if(x1 == ''){x1 = 0;}
    if(x2==''){x2=0;}
    return (-parseFloat(x1) + parseFloat(x2))
    };
    """
    COUNT += 1
    return (str_1 + js_func_3 %format_tuple).replace('back_','back_'+str(COUNT-1))


############################################
#为表格进行相关注释函数
#tag为使用哪一种HTML标签
def  annotation_(filename, title, annotation_1, annotation_2=None):
    string_1 = \
    '\n<p>' + annotation_1 + '</p>'
    
    with open(filename, 'r', encoding='utf8') as f:
        text = f.read()
    match = re.search('title.*?{}</p>.*?subtitle.*?</p>'.format(title), text, re.S).group(0)
    text = text.replace(match, match + string_1, 1)
    if annotation_2 is not None:
        text = text.replace('<th>负载率</th>','<th title="{}">负载率</th>'.format(annotation_2))
    with open(filename, 'w', encoding='utf8') as f:
        f.write(text)

############################################
#对AE分析横向表格yz轴数据进行注入到HTML中，
#解决y轴标签不变的问题
def AE_ylable(timepoint_list, lables_list,filename='designdata.html'):
    ylable_data = """
            "yAxis": [
                        {
                            "show": true,
                            "scale": false,
                            "nameLocation": "end",
                            "nameGap": 15,
                            "gridIndex": 0,
                            "inverse": false,
                            "offset": 0,
                            "splitNumber": 5,
                            "minInterval": 0,
                            "splitLine": {
                                "show": false,
                                "lineStyle": {
                                    "width": 1,
                                    "opacity": 1,
                                    "curveness": 0,
                                    "type": "solid"
                                }
                            },
                            "data":%s
                        }
                    ],

            """
    with open(filename, 'r', encoding='utf8') as f:
        text = f.read()
    #注意这里的index是为了缩小匹配范围，如果调整tab的顺序可能需要调整此处index的值
    index = text.rfind('<td name="物料分析（个）1">')
    prime = text1 = text[index:]
    for timepoint, lables in zip(timepoint_list, lables_list):
        patten = 'AE' + str(timepoint.encode('unicode-escape'))[2:-1]
        str1 =  re.search(r'("title".*?"text".*?'+patten+')',text1, re.S).group()
        first = str1.rfind('"title"')
        text1 = text1.replace(str1[first:], ylable_data %str(lables) + str1[first:],1)
    text = text.replace(prime, text1)
    with open(filename, 'w', encoding='utf8') as f:
        f.write(text)
            

# In[实现table首很固定]
#实现首很固定的python代码
def get_fixed_row(filename='designdata.html'):
    with open(filename,'r', encoding='utf8') as f:
        text = f.read()
    #替换head
    text = text.replace('</head>', string_head + '</head>', 1)
    #创造首行固定的头部
    for id_name in ['产品分析（个）', '物料分析（个）', '绩效分析']:
        patten = re.compile(r'<table id="{}".*?>(.*?</tr>)'.format(id_name),re.S)
        tbody = patten.search(text)
        #去除排序功能
        head_tr = re.sub('onclick="[^"].*?"', '', tbody.group(1))
        head_table = \
        """
        <table class="table" id="head-table_{id_name}" style="display: none;position: fixed;top:0">
        <thead>
        {tbody}
        </thead>
        </table>
        """.format(id_name=id_name, tbody=head_tr)
#        print(head_table + head_tr, '='*50,sep='\n')
        text = text.replace(tbody.group(0), head_table + tbody.group(0), 1)
    
    #替换tab选项卡点击功能
    #固定函数
    fixed_func =[
    "shou_hen('#head-table_产品分析（个）', '#产品分析（个）')",
    "shou_hen('#head-table_物料分析（个）', '#物料分析（个）')",
    "shou_hen('#head-table_绩效分析', '#绩效分析')"
    ]
    
    for tab, func in zip(['产品分析', '物料分析', '绩效分析'], fixed_func):
        patten = '">{}<'.format(tab)
        text = text.replace(patten, " || " + func + patten, 1)
    
    #加入js
    text = text.replace('</html>', js_func_6 + '</html>')
    
    
    with open(filename, 'w', encoding='utf8') as f:
        f.write(text)
     
# In[封装一下对不同表格项的操作]

#为产品分析版块补充功能
def modi_product_analysis():
    annotation_1 = """
    总量 = 目前项目数合计。
    占比 = 目前项目数占去年项目总数的比例
    同比 = 目前项目数同比去年时间点
    存量 = 目前未分配、进行中和待确稿的合计
    人均 = 提案为进行中除以提案设计师人数，签单为进行中除以AE人数
    分配率 = （进行中+已完成+未完成+取消）÷总数
    完成率 =  已完成÷（总数-取消）
    """.strip('\n').replace('\n', '<br />')
    annotation_2 = '负载率=人均 / 满载值'   
    annotation_('designdata.html', '产能分析',annotation_1, annotation_2)        
    prepare_js_fun_2('产品分析（个）')        
    make_up('designdata.html','产品分析（个）',js_func_2)   
    make_up('designdata.html',name=None, js_func=js_func_5)     

#为物料分析版块补充功能
def  modi_material_analysis():
    annotation_1 = \
    """
    完成率：已完成 /（总计 - 取消）
    """.strip()
    length1 = prepare_html('物料分析（个）','designdata.html')
    js_func_wuliao = finish_js_func_3('物料分析（个）',length1)    
    make_up('designdata.html','2020物料分析（个）',js_func_wuliao)
    annotation_('designdata.html', '2020物料分析（个）',annotation_1)

#为产能分析版块补充功能
def modi_productivity_analysis():
     make_up('designdata.html','产能分析',js_func_1)
     
#为绩效分析版块补充功能
def modi_performance_analysis():
    length = prepare_html('绩效分析','designdata.html')
    js_func_jixiao = finish_js_func_3('绩效分析',length)   
    make_up('designdata.html','绩效分析',js_func_jixiao)


#为反馈分析补充的功能
def modi_QR_analysis():
    QR_mediate()



#对html中视觉分析代码进行替换
def vision_html(filename):
    js_string  = """
    <script src="https://cdn.bootcss.com/jquery/1.12.4/jquery.min.js"></script>
    <script src="https://unpkg.com/masonry-layout@4/dist/masonry.pkgd.min.js"></script>
    <script src="https://unpkg.com/imagesloaded@4/imagesloaded.pkgd.js"></script>
    
    <script>
    setTimeout(()=>{$('.grid').imagesLoaded( function() {
    	new Masonry( document.getElementById('container'),{itemSelector:'.grid-item'} );
    });},0);
     </script>
     
     <script>
     $(function(){  
        $(".pimg").click(function(){  
            var _this = $(this);//将当前的pimg元素作为_this传入函数  
            imgShow("#outerdiv", "#innerdiv", "#bigimg", _this);  
        });  
    });  
     </script>
     
     <script>
     function imgShow(outerdiv, innerdiv, bigimg, _this){  
    var src = _this.attr("src");//获取当前点击的pimg元素中的src属性  
    $(bigimg).attr("src", src);//设置#bigimg元素的src属性  
  
        /*获取当前点击图片的真实大小，并显示弹出层及大图*/  
    $("<img/>").attr("src", src).load(function(){  
        var windowW = $(window).width();//获取当前窗口宽度  
        var windowH = $(window).height();//获取当前窗口高度  
        var realWidth = this.width;//获取图片真实宽度  
        var realHeight = this.height;//获取图片真实高度  
        var imgWidth, imgHeight;  
        var scale = 0.8;//缩放尺寸，当图片真实宽度和高度大于窗口宽度和高度时进行缩放  
          
        if(realHeight>windowH*scale) {//判断图片高度  
            imgHeight = windowH*scale;//如大于窗口高度，图片高度进行缩放  
            imgWidth = imgHeight/realHeight*realWidth;//等比例缩放宽度  
            if(imgWidth>windowW*scale) {//如宽度扔大于窗口宽度  
                imgWidth = windowW*scale;//再对宽度进行缩放  
            }  
        } else if(realWidth>windowW*scale) {//如图片高度合适，判断图片宽度  
            imgWidth = windowW*scale;//如大于窗口宽度，图片宽度进行缩放  
                        imgHeight = imgWidth/realWidth*realHeight;//等比例缩放高度  
        } else {//如果图片真实高度和宽度都符合要求，高宽不变  
            imgWidth = realWidth;  
            imgHeight = realHeight;  
        }  
                $(bigimg).css("width",imgWidth);//以最终的宽度对图片缩放  
          
        var w = (windowW-imgWidth)/2;//计算图片与窗口左边距  
        var h = (windowH-imgHeight)/2;//计算图片与窗口上边距  
        $(innerdiv).css({"top":h, "left":w});//设置#innerdiv的top和left属性  
        $(outerdiv).fadeIn("fast");//淡入显示#outerdiv及.pimg  
    });  
      
    $(outerdiv).click(function(){//再次点击淡出消失弹出层  
        $(this).fadeOut("fast");  
    });  
}  
     </script>
    """
    with open(filename, 'r', encoding='utf8') as f:
        text = f.read()
    text = re.sub('<div class="demo">.*?src="".*?</div>',prepare_imgs(), text, 1, re.S)
    text = text.rsplit('</html>',1)[0] +  js_string
    with open(filename, 'w', encoding='utf8') as f:
        f.write(text)
        

# In[主函数]    
def successful():
    modi_product_analysis()
    modi_material_analysis()
    modi_productivity_analysis()
    modi_performance_analysis()
    modi_QR_analysis()
    AE_ylable([name + '数' for name in categories_list], [list(reversed(dataframe.DM.tolist())) for dataframe in categories_df_list],filename='designdata.html')
    get_fixed_row(filename='designdata.html')
    vision_html('designdata.html')
    print('='*50)
    print('{}done!'.format(' '*23))



successful()



