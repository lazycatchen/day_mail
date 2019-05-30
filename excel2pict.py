
import time
import win32com.client as win32
from PIL import ImageGrab
def e2p(pathtable):
        excel=win32.Dispatch('Excel.Application')
        wb=excel.workbooks.open(pathtable)
        wb.Run('日报测试')
        w_pic=wb.worksheets('Sheet1')
        w_pic.Range('B2:F36').CopyPicture()
        w_pic.Paste(w_pic.Range('K1'))
        w_pic.Shapes('Picture 1').copy()
        time.sleep(0.5)
        img1=ImageGrab.grabclipboard()
        pathpict=pathtable[:-9]+'picture.png'
        img1.save(pathpict)
        wb.Close(SaveChanges=0)
        return pathpict


#import Image, ImageFont, ImageDraw

