# -*- coding: utf-8 -*-

#%%
import os
import sys
import time
import datetime
import pandas as pd
import numpy as np
import requests

# 企业微信机器人
webhook = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=4f7adbbc-e23c-4ea0-b450-5a795372c200"
header = {"Content-Type": "text/plain"}

#%%
# 日期
current_time = datetime.datetime.now()
yr, mon, dy = current_time.year, current_time.month, current_time.day
dy_1 = datetime.date.today() - datetime.timedelta(days=1)
# 今年当日
day_gap = int(str(dy_1 - datetime.date(2021, 11, 17))[0:2]) # 今年开售日期
# 去年当日
day_last_year = datetime.date(2020, 12, 3) + datetime.timedelta(days=day_gap)


#%%
# 个人每日投保量统计
count_daily_order_file_path = '/data/users/yeguanhua/SOP/' + str(yr) + '年' + str(mon) + \
                              '月' + str(dy) + '日0时/个人每日投保量统计-截止至' + str(yr) + \
                              '年' + str(mon) + '月' + str(dy) + '日0时.xlsx'
# 如果文件不存在则发送错误消息到企业微信并退出
if not os.path.exists(count_daily_order_file_path):
    content = '《每日投保量统计》文件不存在，无法统计数据。'
    send_data = { "msgtype": "text", "text": { "content": content, "mentioned_list": ["@叶冠华"] } }
    res = requests.post(url = webhook, headers = header, json = send_data)
    sys.exit[0]
else:
    # sheet2:按保司统计-个人订单
    daily_order = pd.read_excel(count_daily_order_file_path, sheet_name='按保司统计-个人订单')
    # 找到合计的行
    daily_order_all = daily_order[daily_order['日期'] == '合计']
    # 找到昨日当日的行
    daily_order_dy_1 = daily_order[daily_order['日期'] == str(dy_1)]

#%%
# 去年个人每日投保量统计
daily_order_last_year_path = '/data/users/yeguanhua/qiye_weixin_bot/个人每日投保量统计-截止至2021年2月1日0时.xlsx'
# 如果文件不存在则发送错误消息到企业微信并退出
if not os.path.exists(daily_order_last_year_path):
    content = '去年的《每日投保量统计》文件不存在，无法统计数据。'
    send_data = { "msgtype": "text", "text": { "content": content, "mentioned_list": ["@叶冠华"] } }
    res = requests.post(url = webhook, headers = header, json = send_data)
    sys.exit[0]
else:
    # sheet2:按保司统计-个人订单
    daily_order_last_year = pd.read_excel(daily_order_last_year_path, sheet_name='按保司统计-个人订单')
    # 找到去年当日的行
    daily_order_last_year = daily_order_last_year[daily_order_last_year['日期'] == str(day_last_year)]

#%%
# 自动扣费及续保情况
count_renew_order_file_path = '/data/users/yeguanhua/SOP/' + str(yr) + '年' + str(mon) + \
                              '月' + str(dy) + '日0时/自动扣费及续保情况-截止至' + str(yr) + \
                              '年' + str(mon) + '月' + str(dy) + '日0时.xlsx'
# 如果文件不存在则直接报错
if not os.path.exists(count_renew_order_file_path):
    content = '《自动扣费及续保情况》文件不存在，无法统计数据。'
    send_data = { "msgtype": "text", "text": { "content": content, "mentioned_list": ["@叶冠华"] } }
    res = requests.post(url = webhook, headers = header, json = send_data)
    sys.exit[0]
else:
    # sheet4：续保情况-今年
    renew_order = pd.read_excel(count_renew_order_file_path, sheet_name='续保情况-今年')
    # 将个人订单续保和企业订单续保合并
    #renew_order['企业订单续保数'] = renew_order['企业订单续保数'].replace(np.nan, 0)
    renew_order['续保总数'] = renew_order['个人订单续保数']# + renew_order['企业订单续保数']
    # 添加新保数量
    #renew_order = renew_order.merge(daily_order_all[['保险公司', '保司净投保量']], how='left')
    #renew_order['保司净投保量'] = renew_order['保司净投保量'].replace(np.nan, 0)
    #renew_order.loc[renew_order['保险公司'] == '合计', '保司净投保量'] = sum(renew_order['保司净投保量'])
    renew_order['订单总数'] = renew_order['今年个人订单总数']# + renew_order['今年企业订单总数']
    renew_order['新保'] = renew_order['订单总数'] - renew_order['续保总数']
    # sheet3：续保情况-去年
    renew_order_last_year = pd.read_excel(count_renew_order_file_path, sheet_name='续保情况-去年')

#%%
content = "每日数据统计（截止{}月{}日0时）\n---\n".format(str(mon), str(dy)) + \
          "**【总体完成{}人】**\n".format(str(sum(daily_order_all['保司净投保量']))) + \
          "思派{}人，".format(daily_order_all.loc[daily_order_all['保险公司'] == '线上平台', '保司净投保量'].iloc[0]) + \
          "人保{}人，".format(daily_order_all.loc[daily_order_all['保险公司'] == '人保财险', '保司净投保量'].iloc[0]) + \
          "国寿{}人，".format(daily_order_all.loc[daily_order_all['保险公司'] == '中国人寿', '保司净投保量'].iloc[0]) + \
          "平安{}人\n---\n".format(daily_order_all.loc[daily_order_all['保险公司'] == '平安养老', '保司净投保量'].iloc[0]) + \
          "**【昨日当日完成{}人】**\n".format(str(sum(daily_order_dy_1['保司净投保量']))) + \
          "思派{}人，".format(daily_order_dy_1.loc[daily_order_dy_1['保险公司'] == '线上平台', '保司净投保量'].iloc[0]) + \
          "人保{}人，".format(daily_order_dy_1.loc[daily_order_dy_1['保险公司'] == '人保财险', '保司净投保量'].iloc[0]) + \
          "国寿{}人，".format(daily_order_dy_1.loc[daily_order_dy_1['保险公司'] == '中国人寿', '保司净投保量'].iloc[0]) + \
          "平安{}人\n---\n".format(daily_order_dy_1.loc[daily_order_dy_1['保险公司'] == '平安养老', '保司净投保量'].iloc[0]) + \
          "**【去年当日完成{}人】**\n".format(str(sum(daily_order_last_year['保司净投保量']))) + \
          "思派去年当日{}人，".format(daily_order_last_year.loc[daily_order_last_year['保险公司'] == '线上平台', '保司净投保量'].iloc[0]) + \
          "人保去年当日{}人，".format(daily_order_last_year.loc[daily_order_last_year['保险公司'] == '人保财险', '保司净投保量'].iloc[0]) + \
          "国寿去年当日{}人，".format(daily_order_last_year.loc[daily_order_last_year['保险公司'] == '中国人寿', '保司净投保量'].iloc[0]) + \
          "平安去年当日{}人\n---\n".format(daily_order_last_year.loc[daily_order_last_year['保险公司'] == '平安养老', '保司净投保量'].iloc[0] + \
                                    daily_order_last_year.loc[daily_order_last_year['保险公司'] == '平安人寿', '保司净投保量'].iloc[0]) + \
          "**【总体新保{}人，".format(renew_order.loc[renew_order['保险公司'] == '合计', '新保'].iloc[0]) + \
          "连续参保{}人】**\n".format(renew_order.loc[renew_order['保险公司'] == '合计', '续保总数'].iloc[0]) + \
          "思派新保{}人，".format(renew_order.loc[renew_order['保险公司'] == '线上平台', '新保'].iloc[0]) + \
          "连续参保{}人\n".format(renew_order.loc[renew_order['保险公司'] == '线上平台', '续保总数'].iloc[0]) + \
          "人保新保{}人，".format(renew_order.loc[renew_order['保险公司'] == '人保财险', '新保'].iloc[0]) + \
          "连续参保{}人\n".format(renew_order.loc[renew_order['保险公司'] == '人保财险', '续保总数'].iloc[0]) + \
          "国寿新保{}人，".format(renew_order.loc[renew_order['保险公司'] == '中国人寿', '新保'].iloc[0]) + \
          "连续参保{}人\n".format(renew_order.loc[renew_order['保险公司'] == '中国人寿', '续保总数'].iloc[0]) + \
          "平安新保{}人，".format(renew_order.loc[renew_order['保险公司'] == '平安养老', '新保'].iloc[0]) + \
          "连续参保{}人\n---\n".format(renew_order.loc[renew_order['保险公司'] == '平安养老', '续保总数'].iloc[0]) + \
          "**【去年累计{}人，".format(renew_order_last_year.loc[renew_order_last_year['保险公司'] == '合计', '去年个人订单总数'].iloc[0]) + \
          "今年累计{}人】**\n".format(renew_order.loc[renew_order['保险公司'] == '合计', '今年个人订单总数'].iloc[0]) + \
          "思派去年累计{}人，".format(renew_order_last_year.loc[renew_order_last_year['保险公司'] == '线上平台', '去年个人订单总数'].iloc[0]) + \
          "今年累计{}人\n".format(renew_order.loc[renew_order['保险公司'] == '线上平台', '今年个人订单总数'].iloc[0]) + \
          "人保去年累计{}人，".format(renew_order_last_year.loc[renew_order_last_year['保险公司'] == '人保财险', '去年个人订单总数'].iloc[0]) + \
          "今年累计{}人\n".format(renew_order.loc[renew_order['保险公司'] == '人保财险', '今年个人订单总数'].iloc[0]) + \
          "国寿去年累计{}人，".format(renew_order_last_year.loc[renew_order_last_year['保险公司'] == '中国人寿', '去年个人订单总数'].iloc[0]) + \
          "今年累计{}人\n".format(renew_order.loc[renew_order['保险公司'] == '中国人寿', '今年个人订单总数'].iloc[0]) + \
          "平安去年累计{}人，".format(renew_order_last_year.loc[renew_order_last_year['保险公司'] == '平安人寿', '去年个人订单总数'].iloc[0] + \
                                    renew_order_last_year.loc[renew_order_last_year['保险公司'] == '平安养老', '去年个人订单总数'].iloc[0]) + \
          "今年累计{}人".format(renew_order.loc[renew_order['保险公司'] == '平安养老', '今年个人订单总数'].iloc[0])

#%%
send_data = { "msgtype": "markdown", "markdown": { "content": content } }
res = requests.post(url = webhook, headers = header, json = send_data)
#print(res.text)
