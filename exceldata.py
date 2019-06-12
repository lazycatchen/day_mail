# _*_ coding: utf-8 _*_
import xlrd
import os
import xlwt
from xlrd import open_workbook
from xlutils.copy import copy

def forder(filename): #返回字典键
    new_flist=[]
    chx=''
    temp=['内蒙','靖边','干北','诺木洪','新疆','河北','江苏','敦煌','共和','山东','尧生']
    for ch in temp:
         if  ch in filename:
             new_flist=temp.index(ch)
             chx=ch
    return chx

def custom_key(word):
   numbers = []
   orderstr=[ '内蒙', '靖边','干北','诺木洪','新疆','河北','江苏','敦煌','共和','山东','尧生']
   temp=[i for i, word1 in enumerate(orderstr) if word1 in word]
   numbers.append(temp)
   return numbers

def exceltable(path):
    g = os.walk(path)
    #pypath='F:\\0\\py_ribao\\py_save\\py日报模板.xls'
    #wb = xlrd.open_workbook(pypath)
    #ws = wb.sheet_by_name('Sheet1')   #修改日报模板，生产日报
    rb = open_workbook('F:\\0\\py_ribao\\py_save\\py日报模板.xls', formatting_info=True)
    wb1 = copy(rb)
    ws = wb1.get_sheet(0)
    #ws.write(0, 0, 'changed!')
    #wb1.save('F:\\0\\py_ribao\\py_save\\py日报模板.xls')

    pathtable='F:\\0\\py_ribao\\py_save\\table.xls'

    for path,dir_list,file_list in g:
        ##########排序!
        file_list.sort(key=custom_key)
        ##
        listdata=[]
        fileorder=[]
        dictdata={}
        xls_result= xlwt.Workbook()
        sht1 = xls_result.add_sheet('Sheet1',cell_overwrite_ok=True)
        j=0

        if int(path[-2:])<10:
            nowdate=int(path[-1])-1
        else:
            nowdate=int(path[-2:])-1
            #
        for file_name in file_list:
            #fileorder.append(forder(file_name))
            data = xlrd.open_workbook(os.path.join(path, file_name))
            names = data.sheet_names()
            table = data.sheets()[len(names)-1]
            ch=forder(file_name)
            #

            if ch=='新疆':
                table=data.sheets()[len(names)-2]
            if ch=='诺木洪' or ch=='共和':
                date1=str(nowdate)
                table=data.sheet_by_name(date1)
            if ch=='靖边':
                date1=str(nowdate)+'日'
                #table=data.sheet_by_name(date1)##### error1
                table=data.sheets()[nowdate-1]

    #for temp in rang(1:15_):
            templist=[]
            for (temp,temp2) in zip(table._cell_types ,table._cell_values) :
                if temp.count(2)==10 or temp.count(2)==11 or temp.count(2)==13:
                    j=j+1   #excel对应的行
                    listdata.append(temp2)
                    for i,data in enumerate(temp2):
                        sht1.write(j,0,ch)
                        ws.write(j,0,ch)#
                        sht1.write(j,i+1,data)

                        ws.write(j,i+1,data)#
                        templist.append(temp2) #单表数据提取
        #datessa=[int(path[-1])-1,int(path[-2:])]
        sht1.write(0,0,int(path[-3]))
        sht1.write(0,1,int(path[-2:])-1)
        #templist.append(datessa)
            #dictdata[ch]=templist  #结果存字典
        xls_result.save(pathtable)
        #ws.save(pathtable)
        #ws.write(0, 0, 'changed!')
        #wb1.sav)
        pathtable1='F:\\0\\py_ribao\\py_save\\'+str(nowdate)+'日'
        return pathtable1



