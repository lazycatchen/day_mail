# _*_ coding: utf-8 _*_
import xlrd
import numpy as np
import types
import pysnooper
import os
import xlwt
@pysnooper.snoop()
def forder(filename): #返回字典键
    new_flist=[]
    temp=['内蒙','靖边','干北','诺木洪','新疆','河北','江苏','敦煌','共和','山东','尧生']
    for ch in temp:
         if  ch in filename:
             new_flist=temp.index(ch)
             ch1=ch
    return ch1

g = os.walk(r"F:\0\4.14")
for path,dir_list,file_list in g:
    listdata=[]
    fileorder=[]
    dictdata={}
    xls_result= xlwt.Workbook()
    sht1 = xls_result.add_sheet('Sheet1',cell_overwrite_ok=True)
    j=0
    for file_name in file_list:
        #fileorder.append(forder(file_name))
        data = xlrd.open_workbook(os.path.join(path, file_name))
        names = data.sheet_names()
        table = data.sheets()[len(names)-1]
        ch=forder(file_name)
        if ch=='新疆':
            table=data.sheets()[len(names)-2]
        if ch=='诺木洪' or ch=='共和':
            table=data.sheets()[13]

#for temp in rang(1:15_):
        templist=[]
        for (temp,temp2) in zip(table._cell_types ,table._cell_values) :
            if temp.count(2)==11 or temp.count(2)==13:
                j=j+1   #excel对应的行
                listdata.append(temp2)
                for i,data in enumerate(temp2):
                    sht1.write(j,0,ch)
                    sht1.write(j,i+1,data)
                    templist.append(temp2) #单表数据提取
        #dictdata[ch]=templist  #结果存字典
    xls_result.save(r"F:\0\new_ribao\table.xlsx")



