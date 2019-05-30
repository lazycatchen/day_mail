# _*_ coding: utf-8 _*_
import win32com.client as wc
pypath='F:\\0\\py_ribao\\20190502\\table - 副本 (2).xls'
app = wc.Dispatch('Excel.Application')
xls = app.WorkBooks.Open(pypath)
app.Run('日报测试')

