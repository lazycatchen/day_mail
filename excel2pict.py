# _*_ coding: utf-8 _*_
# -*- coding: utf-8 -*-

#import os
import time
import win32com.client as win32
from PIL import ImageGrab
#g = os.walk(r"F:\0\new_ribao")
excel=win32.DispatchEx('Excel.Application')
wb=excel.workbooks.open('F:\\0\\new_ribao\\日报.xlsx')
w_pic=wb.worksheets('Sheet1')
w_pic.Range('B2:F36').CopyPicture()
w_pic.Paste(w_pic.Range('K1'))
#w_pic.Shapes('Picture 1').copy()
#w_pic.SaveAs('F:\\0\\new_ribao\\copy.xlsx')
#new_shape_name = 'luzaofa'
#w_pic.Selection.ShapeRange.Name = new_shape_name
w_pic.Shapes('Picture 1').copy()
time.sleep(0.5)
img1=ImageGrab.grabclipboard()
img1.save('F:\\0\\new_ribao\\Picture 1.png')
wb.Close(SaveChanges=0)


#import Image, ImageFont, ImageDraw

