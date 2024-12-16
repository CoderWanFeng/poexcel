# -*- coding: UTF-8 -*-
'''
@学习网站      ：https://www.python-office.com
@读者群     ：http://www.python4office.cn/wechat-group/
@作者  ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫，微信：CoderWanFeng
@代码日期    ：2024/3/20 22:55
'''

import pandas as pd

# 从示例.xlsx文件中读取数据
df = pd.read_excel(r'./示例.xlsx')
left_name=df['姓名'][0]  # 获取左侧合并列的名称
left_c=[]  # 初始化左侧列的值列表
right_c=set()  # 初始化右侧列的值集合
right_l = []  # 初始化存储右侧列名的列表
# 循环处理每列数据，将每列的首个值添加到right_l列表中
for c_name in ['第一列','第二列','第三列','第四列']:
    right_l.append(df[c_name][0])
# 注释掉的代码意图是合并right_c集合，但实际未使用
for l_name in right_l:
    temp_l = str(l_name).split('、')  # 按'、'分割字符串
    for t_name in temp_l:
        if t_name not in right_c:
            right_c.add(t_name)  # 将唯一值添加到right_c集合中
            left_c.append(left_name)  # 将左侧合并列的名称添加到left_c列表中
print(list(right_c))  # 打印右侧列的值集合
print(left_c)  # 打印左侧列的值列表

df1 = pd.DataFrame(left_c)  # 将left_c转换为DataFrame
df2 = pd.DataFrame(list(right_c))  # 将right_c转换为DataFrame
# 根据索引合并df1和df2
result = pd.merge(df1, df2, left_index=True, right_index=True)

# 将合并结果保存到结果.xlsx文件中
result.to_excel(r'./结果.xlsx',index=False)

