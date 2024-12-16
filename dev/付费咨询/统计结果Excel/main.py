# -*- coding: UTF-8 -*-
'''
@学习网站      ：https://www.python-office.com
@读者群     ：http://www.python4office.cn/wechat-group/
@作者  ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫，微信：CoderWanFeng
@代码日期    ：2024/3/15 1:01 
'''

import pandas as pd


def get_index(df: pd.DataFrame):
    return df.columns.tolist()


def get_values(df: pd.DataFrame):
    return df.values.tolist()


def change_content(content: list):
    name_list = content[1].split('、')
    job_list = content[2].split('\n')
    for job_index in range(len(job_list)):
        for name_index in range(len(name_list)):
            if name_list[name_index] in job_list[job_index]:
                job_list[job_index] = job_list[job_index] + '+1'
            else:
                job_list[job_index] = job_list[job_index] + '+0'

    for j_i in range(len(job_list)):
        res_sum = 0
        for i in job_list[j_i].split('+')[1:]:
            res_sum += int(i)

        if res_sum > 0:
            job_list[j_i] = job_list[j_i].split('+')[0] + ':推荐'
        else:
            job_list[j_i] = job_list[j_i].split('+')[0] + ':不推荐'
    res_content = [content[0], content[1], '\n'.join(job_list), content[3]]
    return res_content


if __name__ == '__main__':
    origin_df = pd.read_excel('origin.xlsx')
    index = get_index(origin_df)
    values = get_values(origin_df)
    output_df = pd.DataFrame()
    output_values = []
    for v in values:
        output_values.append(change_content(v))
    df = pd.DataFrame(output_values, columns=index)
    df.to_excel(r'output.xlsx', header=True)
    print(df)
