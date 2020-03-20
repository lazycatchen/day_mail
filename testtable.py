# _*_ coding: utf-8 _*_
import excel2pict
import exceldata
import time
import search_Emax
#datestr = input('请输入起始日期(如20190401): ')   #接受邮件的日期
datestr=str(20200320)
#daydate = input('请输入制表日期(如20190401): ')
for daydate in range(20200305,20200322,1):
    print(daydate)
    datetable ='F:\\0\\py_ribao\\py_save\\'+str(daydate)
    path='F:\\0\\py_ribao\\'+datestr  #存储并以日期命名文件夹，不存在则创建文件夹
    pathtable=exceldata.exceltable(path,datetable)   #完成之后制作日报
    picture=excel2pict.e2p(pathtable)
    time.sleep(3)

#maxele=search_Emax.maxnum()



