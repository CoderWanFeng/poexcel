import pandas as pd

# 读取 JSON 文件
data = pd.read_json('jsfile.json')

# 将数据写入 Excel 文件
data.to_excel('data.xlsx', index=False)
