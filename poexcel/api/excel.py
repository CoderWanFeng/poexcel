#!/usr/bin/env python
# -*- coding:utf-8 -*-
'''
@学习网站      ：https://www.python-office.com
@读者群     ：http://www.python4office.cn/wechat-group/
@作者  ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫，微信：CoderWanFeng
@代码日期    ：2024/1/9 22:56
@本段代码的视频说明     ：
'''
from poexcel.core import QueryExcel
#############################################
# File Name: excel.py
# Mail: 1957875073@qq.com
# Created Time:  2022-4-25 10:17:34
# Description: 有关 excel 的自动化操作
#############################################
from poexcel.core.ExcelType import MainExcel

mainExcel = MainExcel()


# todo:输出文件路径
# @except_dec()
def fake2excel(columns=['name'], rows=1, path='./fake2excel.xlsx', language='zh_CN'):
    """
    视频：https://www.bilibili.com/video/BV1wr4y1b7uk/
    演示代码：
    :param columns:
    :param rows:
    :param path:
    :param language:
    :return:
    """
    mainExcel.fake2excel(columns, rows, path, language)


# 多个excel，合并到一个excel的不同sheet中
# @except_dec()
def merge2excel(dir_path, output_file='merge2excel.xlsx'):
    """
    视频：https://www.bilibili.com/video/BV1Th4y1Y7kd/
    功能：把多个Excel，汇总到1个Excel的不同sheet里面
    演示代码：
    :param dir_path:
    :param output_file:
    :return:
    """
    mainExcel.merge2excel(dir_path, output_file)


# 同一个excel里的不同sheet，拆分为不同的excel文件
# @except_dec()
def sheet2excel(file_path, output_path='./'):
    """
    视频：https://www.bilibili.com/video/BV1714y147Ao/
    演示代码：
    :param file_path:
    :param output_path:
    :return:
    """
    mainExcel.sheet2excel(file_path, output_path)


# 搜索excel中指定内容的文件、行数、内容详情
# PR内容 & 作者：https://gitee.com/CoderWanFeng/python-office/pulls/10
# @except_dec()
def find_excel_data(search_key: str, target_dir: str):
    """
    视频：https://www.bilibili.com/video/BV1Bd4y1B7yr/
    演示代码：

    :param search_key:
    :param target_dir:
    :return:
    """
    mainExcel.find_excel_data(search_key, target_dir)


def excel2pdf(excel_path, pdf_path, sheet_id: int = 0):
    """
    将指定的Excel文件的指定工作表转换为PDF文件。
    视频：https://www.bilibili.com/video/BV1Nh4y1K7KG

    :param excel_path: str, Excel文件的路径。
    :param pdf_path: str, 转换后生成的PDF文件的路径。
    :param sheet_id: int, 工作表的索引，默认为0，表示第一个工作表。
    :return: None
    """
    mainExcel.excel2pdf(excel_path, pdf_path, sheet_id)


def query4excel(query_content, query_path, output_path=None, output_name='output_path/query4excel.xlsx'):
    """
    视频：https://www.bilibili.com/video/BV1Hs4y1S7TT/
    演示代码：

    :param query_content:
    :param query_path:
    :param output_path:
    :param output_name:
    :return:
    """
    QueryExcel.query4excel(query_content, query_path, output_path, output_name)


def count4page(input_path):
    """
    文档：https://blog.csdn.net/weixin_42321517/article/details/131218163
    演示代码：

    :param input_path:
    :return:
    """
    mainExcel.count4page(input_path)


# @except_dec()
def merge2sheet(dir_path, output_sheet_name: str = None, output_excel_name: str = 'merge2sheet'):
    """
    视频：
    演示代码：

    :param dir_path:
    :param output_sheet_name:
    :param output_excel_name:
    :return:
    """
    mainExcel.merge2sheet(dir_path=dir_path, output_excel_name=output_excel_name, output_sheet_name=output_sheet_name)

    # 按指定列的内容，拆分excel
    # PR内容 & 作者：：https://gitee.com/CoderWanFeng/python-office/pulls/11
    # @except_dec()
    def split_excel_by_column(filepath: str, column: int, worksheet_name: str = None):
        """
        视频：
        演示代码：
        :param filepath: 必填，Excel文件的位置和名称
        :param column: 必填，根据第几列拆分
        :param worksheet_name: 选填，可以指定拆分哪一个sheet，不填则默认第一个
        :return:
        """
        mainExcel.split_excel_by_column(filepath, column, worksheet_name)

