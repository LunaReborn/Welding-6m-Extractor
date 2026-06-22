import pandas as pd
import random
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# -------------------------- 配置参数 --------------------------
INPUT_FILE = "1_6m_terms.cleaned.v2.xlsx"         # 你的原始Excel文件名
OUTPUT_FILE = "sampled_2000_data.xlsx" # 输出的抽样Excel文件名
SAMPLE_NUM = 3000                       # 抽样数量

# --- 完全复刻你提供的样式配置 ---
HEADERS = ["术语", "类别", "归一化术语", "原文片段", "来源文档", "文本段", "置信度", "说明"]
TARGET_WIDTHS = [22, 8, 22, 60, 28, 10, 10, 32]
# ---------------------------------------------------------------------------

# 1. 读取Excel数据
try:
    df = pd.read_excel(INPUT_FILE, engine="openpyxl")
    print(f"✅ 成功读取原始数据，总条数：{len(df)}")
except Exception as e:
    print(f"❌ 读取文件失败：{e}")
    exit()

# 2. 随机抽样
if len(df) > SAMPLE_NUM:
    df_sample = df.sample(n=SAMPLE_NUM, random_state=42, replace=False)
    print(f"✅ 已随机抽取 {SAMPLE_NUM} 条数据")
else:
    df_sample = df.copy()
    print(f"✅ 数据总条数不足{SAMPLE_NUM}，已全量导出")

# 3. 重新创建Workbook（而不是直接保存DataFrame，以便完全控制样式）
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "6M术语"

# 写入表头
ws.append(HEADERS)

# 写入数据行
# 确保列顺序一致，按HEADERS顺序取数据
for _, row in df_sample.iterrows():
    data_row = []
    for field in HEADERS:
        # 如果字段不存在，填空字符串
        val = row.get(field, "")
        # 处理NaN等空值
        if pd.isna(val):
            val = ""
        data_row.append(val)
    ws.append(data_row)

# 4. 应用所有样式（完全照搬你的代码逻辑）
header_fill = PatternFill("solid", fgColor="D9EAF7")
header_font = Font(bold=True, color="1F1F1F")
thin = Side(style="thin", color="D9D9D9")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

# 设置表头样式
for cell in ws[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = border

# 设置数据行样式
for row in ws.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        cell.border = border

# 设置列宽
for idx, width in enumerate(TARGET_WIDTHS, start=1):
    ws.column_dimensions[get_column_letter(idx)].width = width

# 冻结窗格和自动筛选
ws.freeze_panes = "A2"
ws.auto_filter.ref = ws.dimensions

# 5. 保存文件
try:
    wb.save(OUTPUT_FILE)
    print(f"✅ 文件已保存为：{OUTPUT_FILE}")
    print(f"📊 最终导出条数：{len(df_sample)}")
except Exception as e:
    print(f"❌ 保存文件失败：{e}")