# coding=utf-8
import schedule
from wxpy import *
import time
# 初始化机器人，扫码登陆
bot=Bot(cache_path=True)
#bot = Bot(console_qr=True, cache_path=True)
myself = bot.self
my_friend = bot.friends().search(u'ssss')[0]
my_friend.send('哈喽')
def job():
    bot.file_helper.send('哈喽 World!！')
schedule.every().day.at("11:06").do(job)
while True:
    schedule.run_pending()#确保schedule一直运行
    time.sleep(1)
bot.join() #保证上述代码持续运行