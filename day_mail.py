# _*_ coding: utf-8 _*_

import poplib
import email
import os
import time
import schedule
import exceldata
import excel2pict
import search_Emax
from wxpy import *
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr

def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        if charset == 'gb2312':
            charset = 'gb18030'
        value = value.decode(charset)
    return value

def get_email_headers(msg):
    headers = {}
    for header in ['From', 'To', 'Cc', 'Subject', 'Date']:
        value = msg.get(header, '')
        if value:
            if header == 'Date':
                headers['Date'] = value
            if header == 'Subject':
                subject = decode_str(value)
                headers['Subject'] = subject
            if header == 'From':
                hdr, addr = parseaddr(value)
                name = decode_str(hdr)
                from_addr = u'%s <%s>' % (name, addr)
                headers['From'] = from_addr
            if header == 'To':
                all_cc = value.split(',')
                to = []
                for x in all_cc:
                    hdr, addr = parseaddr(x)
                    name = decode_str(hdr)
                    to_addr = u'%s <%s>' % (name, addr)
                    to.append(to_addr)
                headers['To'] = ','.join(to)
            if header == 'Cc':
                all_cc = value.split(',')
                cc = []
                for x in all_cc:
                    hdr, addr = parseaddr(x)
                    name = decode_str(hdr)
                    cc_addr = u'%s <%s>' % (name, addr)
                    cc.append(to_addr)
                headers['Cc'] = ','.join(cc)
    return headers

def get_email_content(message, savepath,datestr):
    attachments = []
    for part in message.walk():
        filename = part.get_filename()
        if filename:
           date1  = time.strptime(message.get("Date")[0:24],'%a, %d %b %Y %H:%M:%S') #格式化收件时间
           date2 = time.strftime("%Y%m%d", date1)
           # if date2!=datestr:    #寻找指定日期的邮件并下载
           if date2>datestr:
              break
           else:
               try:
                     filename = decode_str(filename)
                     data = part.get_payload(decode=True)
                     abs_filename = os.path.join(savepath, filename)
                     attach = open(abs_filename, 'wb')
                     attachments.append(filename)
                     attach.write(data)
                     attach.close()
               except Exception as e:
                   print(e)
                   continue
    return attachments


def day_mail():
       email = 'lnxnycsq@163.com'
       # password= input('请输入密码: ')
       password='ln2018'
       pop3_server = 'pop.163.com'
       server = poplib.POP3_SSL(pop3_server)
       # 可以打开或关闭调试信息:
       server.set_debuglevel(0)
       # POP3服务器的欢迎文字:
       print(server.getwelcome())
        # 身份认证:
       server.user(email)
       server.pass_(password)
       # stat()返回邮件数量和占用空间:
       msg_count, msg_size = server.stat()
       print('message count:', msg_count)
       print('message size:', msg_size, 'bytes')
       # b'+OK 237 174238271' list()响应的状态/邮件数量/邮件占用的空间大小
       resp, mails, octets = server.list()
       namelist=[]
       namestr=""
       namelist1=[]

       for i in range(1, msg_count+1):
           resp, byte_lines, octets = server.retr(i)  # 转码
           str_lines = []
           for x in byte_lines:
               str_lines.append(x.decode()) # 拼接邮件内容
           msg_content = '\n'.join(str_lines)  # 把邮件内容解析为Message对象
           msg = Parser().parsestr(msg_content)
           #headers = get_email_headers(msg)
           attachments = get_email_content(msg, path,datestr)
           #print('subject:', headers['Subject'])
           print('attachments: ', attachments)
           namelist.append(attachments)
           namelist1+=attachments

       temp=['内蒙','靖边','干北','诺木洪','新疆','河北','江苏','敦煌','共和','山东','尧生']
       namestr=''.join(namelist1)   #文件夹名字合并
       messagesend=""
       for ch in temp:    #寻找有无未发送文件的公司
            if  ch in namestr:
                my_friend.send(ch+'完成')   #完成的场站
                time.sleep(0.5)
            else:
                messagesend=messagesend+ch+"、"
                time.sleep(0.5)
       if messagesend:
            my_friend.send('请'+messagesend[:-1]+'公司（风电场、光伏电站）尽快报送日报')  #未完成的场站
       else:
          pathtable=exceldata.exceltable(path)   #完成之后制作日报
          picture=excel2pict.e2p(pathtable)   #制作日报
          my_friend.send_image(picture)
          maxele=search_Emax.maxnum()
          my_friend.send(maxele)
          #ribao_groups.send_image(picture)   #将日报转图片发送到微信群
          sys.exit()

if __name__ == '__main__':
    # 账户信息
    bot=Bot(cache_path=True)
    my_friend = bot.friends().search(u'ssss')[0]  #寻找微信-收件人名字
    #ribao_groups = bot.groups().search(u'日报')[0]
    my_friend.send('test')
    #ribao_groups.send('test')
    datestr = input('请输入起始日期(如20190401): ')   #接受邮件的日期

    path='F:\\0\\py_ribao\\'+datestr  #存储并以日期命名文件夹，不存在则创建文件夹
    isExists=os.path.exists(path)
    if not isExists:
        os.makedirs(path)
    else:
        print(path+' 目录已存在')
    day_mail()          #执行一次邮件下载程序
    schedule.every(10).minutes.do(day_mail) #每隔三分钟执行一次
    #schedule.every().day.at("09:10").do(day_mail)
    while True:
        schedule.run_pending()#确保schedule一直运行
        time.sleep(1)
    bot.join()

