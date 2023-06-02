# -*- coding: utf-8 -*-
import pandas as pd
import json
def excel_to_json(file_path):
    # 读取Excel表格
    df = pd.read_excel(file_path)
    # 转换为JSON对象数组
    json_data = {"list": []}
    for index, row in df.iterrows():
        json_item = {
            "__id": "9E4B7B7D-64D1-4F63-9255-BCFE7C468971",
            "taskdesc": row["工作内容"],
            "projectid": "PRT16240943860631473631",
            "planbegintime": str(row["计划开始时间"].date()),
            "plancompletime": str(row["计划完成时间"].date()),
            "workload":  row["工作量"],
            "executor": row["责任人"],
            "reviewer":  row["审批人"],
            "sourcedesc": ""
        }
        json_data["list"].append(json_item)
    
    return json_data

# 输入Excel文件路径
file_path = input("请输入Excel文件路径: ")

# 转换为JSON对象数组
json_data = excel_to_json(file_path)

# 输出JSON数据
json_str = json.dumps(json_data, indent=4)

# 保存为文件
output_file = "MMKPI.txt"
with open(output_file, "w",encoding="utf-8") as file:
    file.write(json_str)

print("已将JSON数据保存为文件: " + output_file)