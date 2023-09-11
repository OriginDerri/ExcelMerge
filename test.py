
import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


# 获取文件夹路径
folder_path = "C:\\Users\\Neuvix\\Desktop\\test11\\test111"  # 替换为文件夹路径

# 创建Excel写入器
writer = pd.ExcelWriter('merged_file.xlsx', engine='xlsxwriter')

# 创建Excel工作簿
workbook = writer.book

# 遍历文件夹中的所有CSV文件
for file_name in os.listdir(folder_path):
    if file_name.endswith('.csv'):
        file_path = os.path.join(folder_path, file_name)
        sheet_name = os.path.splitext(file_name)[0]  # 使用文件名作为工作表名

        # 读取CSV文件
        df = pd.read_csv(file_path,dtype=str)

        # 将CSV文件写入工作表
        df.to_excel(writer, sheet_name=sheet_name, index=False)

# 保存Excel文件
writer.close()

print("合并完成！")