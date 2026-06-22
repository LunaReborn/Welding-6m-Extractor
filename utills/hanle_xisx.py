import pandas as pd

input_file = "sampled_3000_data.xlsx"
output_file = "sampled_3000_data_matched.xlsx"

# 读取 Excel
df = pd.read_excel(input_file, sheet_name="6M术语")

# 1. 删除“术语”中字数超过 7 的行
df = df[df["术语"].astype(str).str.len() <= 7]

# 2. 删除指定列
drop_cols = ["原文片段", "来源文档", "文本段"]
df = df.drop(columns=[col for col in drop_cols if col in df.columns])

# 3. 按 5E1M / 6M 文档规则校验“术语”和“类别”
category_keywords = {
    "人": ["焊工", "检验员", "持证焊工", "作业人员", "监护人", "人员", "岗位", "资质"],
    "机": ["焊机", "焊枪", "送丝机", "夹具", "烘干箱", "设备", "工具", "工装", "仪器"],
    "料": ["焊丝", "母材", "不锈钢", "氩气", "钨极", "焊剂", "合金", "钢", "气体", "材料"],
    "法": ["焊", "加热", "电流", "温度", "坡口", "工艺", "方法", "参数", "规范"],
    "环": ["通风", "湿度", "风速", "作业区域", "防火", "环境", "现场"],
    "测": ["检查", "检测", "检验", "试验", "尺寸", "游标卡尺", "测量", "质量"]
}

def is_match(term, category):
    term = str(term)
    category = str(category)

    # 如果类别不在规则中，删除
    if category not in category_keywords:
        return False

    # 术语命中当前类别关键词，保留
    for kw in category_keywords[category]:
        if kw in term:
            return True

    # 如果术语命中了其他类别关键词，则删除
    for other_cat, keywords in category_keywords.items():
        if other_cat == category:
            continue
        for kw in keywords:
            if kw in term:
                return False

    # 无明显冲突时默认保留
    return True

df = df[df.apply(lambda row: is_match(row["术语"], row["类别"]), axis=1)]

# 4. 保存结果
df.to_excel(output_file, index=False)

print(f"处理完成，已保存为：{output_file}")