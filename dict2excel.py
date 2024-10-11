import pandas as pd

user_dict = {"姓名": "张三", "年龄": "33", "性别": "男"}

# 将字典转换为数据框
df = pd.DataFrame([user_dict])

# 将数据框写入 Excel 文件
df.to_excel('user_data.xlsx', index=False)