# -*- coding: utf-8 -*-

#%%
import imaplib
import email
from email.header import decode_header
import os
import sys
import datetime
import pandas as pd
import numpy as np
import requests
#from dateutil.relativedelta import relativedelta
# Get current datetime
current_time = datetime.datetime.now()
yr, mon, dy, hr = current_time.year, current_time.month, current_time.day, current_time.hour
# 如果运行时间为早上6点前，一律检测0点的邮件
if int(hr) < 6: 
    hr = '00'
# 邮件发送时间
sent_time = str(yr) + '年' + str(mon) + '月' + str(dy) + '日0时'

#%%
### 自定义信息 ###
# 项目名称
pname = 'project_name'
# 最后一封发送的每日参保量报表邮件的标题
email_subject = pname + '统计数据-截止至' + sent_time
# 企业微信机器人
webhook = "https://qyapi.weixin.qq.com/webhook/"
header = {"Content-Type": "application/json"}
# 连接邮箱
conn=imaplib.IMAP4_SSL(port="993", host="imap.exmail.qq.com")
conn.login("your_email@email.com", "dynamic_passwd")

#%%
# 显示邮箱所有文件夹
#print(conn.list())
# 选择文件夹
conn.select('"Sent Messages"')
# 获取邮件（经测试搜索功能在腾讯企业邮箱无法使用，所以直接获取全部邮件）
resp, messages = conn.search(None, "ALL")
# 倒序遍历前50封邮件
x = 0
email_sent = 0
for num in reversed(messages[0].split()):
    x = x + 1
    if x > 50:
        break
    resp, data = conn.fetch(num, '(RFC822)')
    # Skip [NoneType] list data
    if data == [None]:
        continue
    msg = email.message_from_bytes(data[0][1])
    # 获取邮件标题
    msg_sub = msg.get('subject')
    msg_header = decode_header(msg_sub)[0][0].decode('utf-8')
    # 如果邮件标题不含有项目名称，则跳过，否则发送记录加一
    if email_subject not in msg_header:
        continue
    else:
        email_sent = email_sent + 1
conn.close()
conn.logout()
# 如果一封邮件都没有发，则在通知发送失败
if email_sent == 0:
    send_data = {
        "msgtype": "text", 
        "text": {
            "content": pname + '-' + sent_time + "数据没有发送成功"
        }
    }
    res = requests.post(url = webhook, headers = header, json = send_data)
# 如果发送的邮件超过一封，则通知有重复发送
elif email_sent > 1:
    send_data = {
        "msgtype": "text", 
        "text": {
            "content": pname + '-' + sent_time + "数据有重复发送"
        }
    }
    res = requests.post(url = webhook, headers = header, json = send_data)
