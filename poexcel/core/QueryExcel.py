# -*- coding: UTF-8 -*-
'''
@作者 ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@微信 ：CoderWanFeng : https://mp.weixin.qq.com/s/yFcocJbfS9Hs375NhE8Gbw
@个人网站 ：www.python-office.com
@Date    ：2023/3/22 20:49 
@Description     ：
'''
import os
from pathlib import Path

import pandas as pd
from poprogress import simple_progress


def query4excel(query_content, query_path, output_path, output_name):

    if not output_name.endswith('xlsx') and not output_name.endswith('xls'):
        print('output_name必须以.xlsx或者.xls结尾')
        return
    abs_query_path = Path(query_path).absolute()
    if output_path:
        abs_output_path = Path(output_path).absolute() / output_name
    else:
        abs_output_path = abs_query_path / output_name
    pwd = abs_output_path.parent
    # t = type(pwd)
    if not pwd.exists():
        pwd.mkdir()
    waiting_query_excel_files = []
    # 如果不存在，则不做处理
    if not abs_query_path.exists():
        print("path does not exist path = " + query_path)
        return
    # 判断是否是文件
    elif abs_query_path.is_file():
        print("path file type is file " + query_path)
        waiting_query_excel_files.append(query_path)
    # 如果是目录，则遍历目录下面的文件
    elif abs_query_path.is_dir():
        for dirpath, dirnames, filenames in simple_progress(os.walk(str(abs_query_path)),
                                                            desc=f'正在查找：{query_path}：'):
            for filename in filenames:
                if filename.endswith('.xlsx') or filename.endswith('.xls'):
                    waiting_query_excel_files.append(Path(dirpath).absolute() / filename)
    res_df = pd.DataFrame()
    print(f'{query_path}下，一共有{str(len(waiting_query_excel_files))}个Excel')
    for excel in simple_progress(waiting_query_excel_files, desc='正在搜索Excel中符合条件的数据：'):
        df = pd.read_excel(excel, sheet_name=None, header=None)
        for sheet in df.values():
            for i, current_row in enumerate(sheet.itertuples()):
                # if i == 0:
                #     print(sheet.iloc[i].to_dict())
                if query_content in current_row:
                    file_dict = {"文件位置":str(excel)}
                    file_dict.update(sheet.iloc[i].to_dict())
                    current_row_df = pd.DataFrame(file_dict, index=[0])
                    res_df = res_df.append(current_row_df)
                    # print()
    # print(res_df)
    res_df.to_excel(str(abs_output_path),index=False)
    print(f'完成搜索，结果存储在：{str(abs_output_path)}')

