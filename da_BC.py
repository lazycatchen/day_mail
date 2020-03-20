import xlrd
import os
import xlwt
from xlrd import open_workbook
from xlutils.copy import copy
data = xlrd.open_workbook('F:\\0\\new_ribao\\报表\\20191231.xls',formatting_info=True)
pathtable='F:\\0\\new_ribao\\报表\\table.xls'

listdata=[]
xls_result= xlwt.Workbook()
sht1 = xls_result.add_sheet('Sheet1',cell_overwrite_ok=True)

templist=[]
j=0

for table in data.sheets():
     ch=table.name
     for (temp,temp2) in zip(table._cell_types ,table._cell_values) :  #读取文件中数据
                if  temp.count(2)==11 :   #根据其独特的数据结构，只要是数字其_cell_types为2
                    j=j+1 #excel对应的行
                    listdata.append(temp2)
                    for i,data in enumerate(temp2):  #遍历存入
                        sht1.write(j,0,ch)
                        #ws.write(j,0,ch)   #
                        sht1.write(j,i+1,data)
                       # ws.write(j,i+1,data)  #
                       # templist.append(temp2) #单表数据提取
xls_result.save(pathtable)