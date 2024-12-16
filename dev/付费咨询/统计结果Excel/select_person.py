# -*- coding: UTF-8 -*-
'''
@学习网站      ：https://www.python-office.com
@读者群     ：http://www.python4office.cn/wechat-group/
@作者  ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫，微信：CoderWanFeng
@代码日期    ：2024/3/17 6:49 
@本段代码的视频说明     ：
'''
import pandas as pd

# 读取候选人和投票结果的Excel文件，示例代码未实际读取，仅展示结构
# pd.read_excel('1.xlsx')  # 候选人
# pd.read_excel('2.xlsx')  # 投票结果

# 全部姓名列表
all_name_list = ["小明", "小李", "小张", "小王"]

# 示例代码，用于读取含有班级和姓名信息的Excel，并初始化worker_data字典
# df = pd.read_excel('select_person.xlsx')

# tea = df["班级"]
# stu = df["姓名"]

# worker_data = {}

# for t_index in range(len(tea)):
#     worker_data[tea[t_index]] = [stu[t_index].split('、')]
#     worker_data[tea[t_index]] +'/n'+ [stu[t_index]]

# 职务与对应的学生列表
m_list = ["张老师", "李老师"]
worker_data = {
    "张老师-主任": ["小王", "小张"],
    "李老师-老师": ["小王"],
}

# 存储不在all_name_list中的姓名
error_name = set()

# 检查worker_data中的学生姓名是否都在all_name_list中
for t, stu in worker_data.items():
    for name in stu:
        if name not in all_name_list:
            error_name.add(name)
print(error_name)

# 如果有错误的姓名，则输出错误名单及其统计信息
if len(error_name) > 0:
    print("错误名单：", error_name)
    print("错误名单数量：", len(error_name))
    print("错误名单占比：", len(error_name) / len(all_name_list))
    print()
else:
    # 原始姓名列表，示例代码假设已知，无需从Excel读取
    # ori = df["姓名"]
    ori = ["小王", "小张", "小明", "小李"]

    # 初始化结果DataFrame
    df_result = pd.DataFrame()

    # 遍历原始姓名列表，统计每个人在worker_data中的职务数量
    for name in ori:
        manager_num = 0  # 主任职务数量
        worker_num = 0  # 老师职务数量

        for tea_work, stu in worker_data.items():
            if name in stu:
                if "主任" in tea_work:
                    manager_num += 1
                elif "老师" in tea_work:
                    worker_num += 1
                # break

        # 打印统计结果
        print(name, "主任", manager_num, "老师", worker_num)

        # 构建当前人员的统计结果DataFrame
        df_result_current = pd.DataFrame()
        df_result_current["姓名"] = [name]
        df_result_current["主任"] = [manager_num]
        df_result_current["老师"] = [worker_num]
        df_result_current["主任率"] = [manager_num / len(m_list)]

        # 将当前人员的统计结果添加到总结果DataFrame中
        df_result = df_result.append(df_result_current, ignore_index=True)

    # 输出错误名单（如果存在）
    for error_name001 in error_name:
        print(error_name001)
    print(df_result)
    # 将最终统计结果保存到Excel文件中
    df_result.to_excel("select_person_result.xlsx")
