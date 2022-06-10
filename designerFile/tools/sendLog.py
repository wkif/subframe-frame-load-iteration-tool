#! /usr/bin/env python
# coding=utf-8
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from smtplib import SMTP_SSL
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header

from PyQt5.QtWidgets import QMessageBox


def sendLog():
    # qq邮箱smtp服务器
    host_server = 'smtp.163.com'
    # sender_qq为发件人的qq号码
    sender_qq = 'loadtoolsbug@163.com'
    # pwd为qq邮箱的授权码
    pwd = 'IJDSSSEUKWYQODUR'  ##
    # 发件人的邮箱
    sender_qq_mail = 'loadtoolsbug@163.com'
    # 收件人邮箱
    receiver = 'kif101001000@163.com'

    # 邮件的正文内容
    mail_content = "老大，这是客户提交的Bug日志"
    # 邮件标题
    mail_title = '来Bug日志啦！'

    # 邮件正文内容
    msg = MIMEMultipart()
    # msg = MIMEText(mail_content, "plain", 'utf-8')
    msg["Subject"] = Header(mail_title, 'utf-8')
    msg["From"] = sender_qq_mail
    msg["To"] = Header("接收者测试", 'utf-8')  ## 接收者的别名

    # 邮件正文内容
    msg.attach(MIMEText(mail_content, 'html', 'utf-8'))
    dir = './log'

    file_lists = os.listdir(dir)
    file_lists.sort(key=lambda fn: os.path.getmtime(dir + "\\" + fn) if not os.path.isdir(dir + "\\" + fn) else 0)
    file = os.path.join(dir, file_lists[-1])
    # 构造附件1，传送当前目录下的 test.txt 文件

    att1 = MIMEText(open(file, 'rb').read(), 'base64', 'utf-8')
    att1["Content-Type"] = 'application/octet-stream'
    # 这里的filename可以任意写，写什么名字，邮件中显示什么名字
    att1["Content-Disposition"] = 'attachment; filename="bug.txt"'
    msg.attach(att1)

    # ssl登录
    smtp = SMTP_SSL(host_server)
    # set_debuglevel()是用来调试的。参数值为1表示开启调试模式，参数值为0关闭调试模式
    smtp.set_debuglevel(1)
    smtp.ehlo(host_server)
    smtp.login(sender_qq, pwd)

    smtp.sendmail(sender_qq_mail, receiver, msg.as_string())
    smtp.quit()
    msg_box = QMessageBox(QMessageBox.Information,'提示', '发送成功')
    msg_box.exec_()
