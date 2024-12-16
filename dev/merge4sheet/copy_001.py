# -*- coding: UTF-8 -*-
'''
@学习网站      ：https://www.python-office.com
@读者群     ：http://www.python4office.cn/wechat-group/
@作者  ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫，微信：CoderWanFeng
@代码日期    ：2024/3/2 14:20 
@本段代码的视频说明     ：
'''

import os
import pandas as pd

# 指定包含Excel文件的文件夹路径
folder_path = r'D:\workplace\code\github\poexcel\dev\merge4sheet\test_files'

# 获取文件夹中的所有Excel文件
excel_files = [file for file in os.listdir(folder_path) if file.endswith(('.xls', '.xlsx'))]

# 创建一个空的DataFrame来存储合并后的数据
merged_data = pd.DataFrame()

# 遍历所有Excel文件
for file in excel_files:
    file_path = os.path.join(folder_path, file)

    # 读取Excel文件，获取所有工作表
    xls = pd.ExcelFile(file_path)

    # 遍历每个工作表并合并它们
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # 添加一个新的列，用于标识数据来自哪个Excel文件的哪个工作表
        # df['SourceFile'] = file
        # df['SheetName'] = sheet_name

        # 合并数据，将当前工作表的数据追加到已合并的数据中
        merged_data = merged_data._append(df, ignore_index=True)

# 将合并后的数据保存为一个新的Excel文件，指定index=False以避免保存索引列
merged_data.to_excel(r'D:\workplace\code\github\poexcel\dev\merge4sheet\test_files\\合并数据-1.xlsx', index=False)

print('Excel文件合并完成并保存为合并数据.xlsx，包含标识列SourceFile和SheetName')