# _*_ coding: utf-8 _*_
import time
import win32com.client as win32
from PIL import ImageGrab
def e2p(pathpicture):
        excel=win32.Dispatch('Excel.Application')
        pathtable='F:\\0\\py_ribao\\py_save\\py日报模板.xls'  #打开日报模板
        wb=excel.workbooks.open(pathtable)
        excel.Run('日报测试')    #运行vba
        wb.Close(SaveChanges=1)     #保存
        time.sleep(0.1)
        wb=excel.workbooks.open(pathtable)     #打开文件，复制截图
        w_pic=wb.worksheets('Sheet2')
        w_pic.Range('B2:F36').CopyPicture()
        w_pic.Paste(w_pic.Range('K1'))
        w_pic.Shapes('Picture 1').copy()
        time.sleep(0.5)      #给缓存一点时间，防止一闪即逝抓不到图
        img1=ImageGrab.grabclipboard()
        pathpict=pathpicture+'.png'
        img1.save(pathpict)
        wb.Close(SaveChanges=0)
        excel.Quit()
        return pathpict




