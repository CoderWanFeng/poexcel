# -*- coding: UTF-8 -*-
'''
@学习网站      ：https://www.python-office.com
@读者群     ：http://www.python4office.cn/wechat-group/
@作者  ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫，微信：CoderWanFeng
@代码日期    ：2024/12/16 22:44 
@本段代码的视频说明     ：
'''

"""
以下是定制开发
"""
import pandas as pd


def string2excel_20241216(string, output_path='./string2excel.xlsx'):
    # 定义运维工单信息字典
    work_order = {
        "故障内容": "世贸外滩花园 4单元货梯启用梯控和人脸识别设备",
        "对接人员": "保安队长",
        "联系方式": "18720226952",
        "客户名称": "世贸外滩",
        "客户地址": "",  # 客户地址未提供，这里留空
        "项目名称": "U0170",
        "是否加急": False,  # 默认为否
        "业务负责人": "杨开坊",
        "是否收费": False,
        "是否我司供货": True
    }

    # 将字典转换为DataFrame
    df = pd.DataFrame([work_order])

    # 将DataFrame写入Excel文件
    output_file = '运维工单信息.xlsx'
    df.to_excel(output_file, index=False, engine='openpyxl')

    print(f"运维工单信息已保存至 {output_file}")
