# _*_ coding: utf-8 _*_
import xlwt
from xlrd import open_workbook
from xlutils.copy import copy
rb = open_workbook('F:\\0\\py_ribao\\py_save\\py日报模板.xls', formatting_info=True)
xlscopy=copy(rb)
sheet1=xlscopy.getsheet('伊和')
orgin_sheet=open_workbook('F:\\0\\py_ribao\\py_save\\py日报模板.xls', formatting_info=True)
names = orgin_sheet.sheet_names()
table = orgin_sheet.sheets()[len(names)-1]
