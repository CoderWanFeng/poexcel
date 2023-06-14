# -*- coding: UTF-8 -*-
'''
@Author  ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@WeChat     ：CoderWanFeng
@Blog      ：www.python-office.com
@Date    ：2023/6/14 22:31 
@Description     ：
'''

import os
import win32com.client

# 指定文件夹路径
folder_path = r"D:\workplace\code\github\poexcel\dev"

# 打开 Excel 应用程序
excel = win32com.client.Dispatch("Excel.Application")

# 遍历文件夹下的所有文件
for file_name in os.listdir(folder_path):
    # 判断文件是否是 Excel 文件
    if file_name.endswith(".xlsx") or file_name.endswith(".xls"):
        # 打开 Excel 文件
        file_path = os.path.join(folder_path, file_name)
        workbook = excel.Workbooks.Open(file_path)
        # 获取 Excel 文件的打印页数
        page_count = workbook.ActiveSheet.PageSetup.Pages.Count
        # 输出 Excel 文件的打印页数
        print(f"{file_name}: {page_count}页")
        # 关闭 Excel 文件
        workbook.Close()

# 关闭 Excel 应用程序
excel.Quit()
