# -*- coding: UTF-8 -*-
'''
@作者  ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫，微信：CoderWanFeng
@读者群     ：http://www.python4office.cn/wechat-group/
@学习网站      ：https://www.python-office.com
@代码日期    ：2023/8/6 17:16 
@本段代码的视频说明     ：
'''

import poexcel

poexcel.split_excel_by_column(filepath=r'D:\workplace\code\github\python-office\demo\poexcel\excel\split_excel_by_column.xlsx',
                              column=1,
                              worksheet_name='platform')
