# -*- coding: UTF-8 -*-
'''
@Author  ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@WeChat     ：CoderWanFeng
@Blog      ：www.python-office.com
@Date    ：2023/3/25 15:50 
@Description     ：
'''
import poexcel

poexcel.query4excel(query_content='必填，需要查询的内容',
                    query_path=r'必填，放Excel文件的位置',
                    output_path=r'选填，输出查询结果Excel的位置，默认是query_path的位置',
                    output_name='选填，输出的文件名字，默认是：query4excel.xlsx')
