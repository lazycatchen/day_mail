# _*_ coding: utf-8 _*_
import xlrd
import os
import xlwt
from xlrd import open_workbook
def maxnum():
        old_max = open_workbook('F:\\0\\py_ribao\\py_save\\old_max.xls', formatting_info=True)
        data= old_max.sheets()[0]._cell_values
        rb = open_workbook('F:\\0\\py_ribao\\py_save\\table.xls', formatting_info=True)
        tempdata= rb.sheets()[0]._cell_values
        i=0
        xls_result1= xlwt.Workbook()
        sht1 = xls_result1.add_sheet('Sheet1',cell_overwrite_ok=True)
        listi=[]
        listname=["日期","伊和乌素","乌吉尔","杭锦旗","包头","新发","靖边","烟墩山","干三","南北","盐池","都兰1、2","都兰3","多能风电","达坂城","十三间房","小草湖","康保2","康保3","大市","丰宁","东台","\
        敦煌","共和","格尔木1","格尔木2","多能光伏","储能","莒县","枣庄","宜君","光热"] #r1
        for (maxnum,temp,name) in zip(data ,tempdata,listname):
            if maxnum[0]<temp[1]:
                data[i][0]=temp[1]
                str1='恭喜'+name+'取得今年来最高发电量'+str(temp[1])+'万千瓦时'

                listi.append(str1)

            sht1.write(i,0,data[i][0])
            i=i+1

        xls_result1.save('F:\\0\\py_ribao\\py_save\\old_max.xls')
        print(listi)
        return listi