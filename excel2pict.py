
import time
import win32com.client as win32

from PIL import ImageGrab
def e2p(pathpicture):
        excel=win32.Dispatch('Excel.Application')
        pathtable='F:\\0\\py_ribao\\py_save\\py日报模板.xls'
        wb=excel.workbooks.open(pathtable)
        excel.Run('日报测试')
        wb.Close(SaveChanges=1)
        wb=excel.workbooks.open(pathtable)
        w_pic=wb.worksheets('Sheet2')
        w_pic.Range('B2:F36').CopyPicture()
        w_pic.Paste(w_pic.Range('K1'))
        w_pic.Shapes('Picture 1').copy()
        time.sleep(0.5)
        img1=ImageGrab.grabclipboard()
        pathpict=pathpicture+'日.png'
        img1.save(pathpict)
        wb.Close(SaveChanges=0)
        excel.Quit()
        return pathpict


#import Image, ImageFont, ImageDraw

