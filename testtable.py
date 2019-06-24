# _*_ coding: utf-8 _*_
import excel2pict
import exceldata
datestr = input('请输入起始日期(如20190401): ')   #接受邮件的日期

path='F:\\0\\py_ribao\\'+datestr  #存储并以日期命名文件夹，不存在则创建文件夹
pathtable=exceldata.exceltable(path)   #完成之后制作日报
picture=excel2pict.e2p(pathtable)



