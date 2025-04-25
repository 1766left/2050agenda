import pandas as pd
import json

# 加载Excel文件并跳过标题行
df = pd.read_excel('1.xlsx', header=None, skiprows=1)

# 打印前几行数据以便检查
print("Excel数据前3行:")
print(df.head(3))

# 解析每行JSON数据
parsed_data = []
for index, row in df.iloc[:, 0].items():
    try:
        data = json.loads(row)
        parsed_data.append(data)
    except json.JSONDecodeError as e:
        print(f"第 {index+2} 行JSON解析错误: {e}")  # +2是因为跳过了标题行且索引从0开始

# 根据地点和时间段合并数据
merged_data = {}
for item in parsed_data:
    key = (item['地点'], item['时间段'])
    activity = {k: v for k, v in item.items() if k not in ['地点', '时间段']}
    
    if key not in merged_data:
        merged_data[key] = []
    merged_data[key].append(activity)

# 根据序号排序（序号为空或非数字默认9999，放到最后）
final_data = []
for (location, period), activities in merged_data.items():
    sorted_activities = sorted(activities, key=lambda x: int(x.get('序号') or 9999))
    final_data.append({
        "地点": location,
        "时间段": period,
        "活动": sorted_activities
    })

# 转换为JSON格式并保存
with open('merged_activities.json', 'w', encoding='utf-8') as f:
    json.dump(final_data, f, ensure_ascii=False, indent=2)

print("数据已合并并保存到 merged_activities_new.json")
