#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
清洗 Welding-6m-Extractor 导出的 6M/5M1E 术语 Excel。v2

相比 v1 的改进：
1. 按 5M1E 定义做类别边界校验：人/机/料/法/环/测。
2. 不确定项进入 review.xlsx，不再一刀切删除。
3. 料类保留焊材牌号、钢号、焊丝/焊条型号，如 H10MnSi、W707Ni、E4303、ER50-6、Q235。
4. 裸化学式分为必删和待复核：CO、NO 等删除；CO2、Ar、O2、N2 等默认进入复核，可用 --drop-gas-formula 删除。
5. 人名过滤更保守：优先保留岗位/角色词，疑似姓名尽量结合作者、主编、参考文献等上下文。
6. 额外输出 category_suspect.xlsx，提示可能错分的术语。

用法：
    python filter_6m_terms_v2.py 1_6m_terms.xlsx \
      -o 1_6m_terms.cleaned.v2.xlsx \
      --removed 1_6m_terms.removed.v2.xlsx \
      --review 1_6m_terms.review.v2.xlsx \
      --category-suspect 1_6m_terms.category_suspect.v2.xlsx

可选：
    --drop-gas-formula    将 CO2、Ar、O2、N2、H2、He 等裸气体化学式也删除，而不是进入复核。
    --dedupe-by-doc-chunk 去重时保留同一术语在不同来源文档/文本段中的记录。

依赖：
    pip install openpyxl
"""

from __future__ import annotations

import argparse
import re
from collections import OrderedDict
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ALLOWED_CATEGORIES = {"人", "机", "料", "法", "环", "测"}

# 5M1E 类别边界：用于提示错分，不直接删除。
CATEGORY_HINTS = {
    "人": {
        "操作者", "操作人员", "作业人员", "施工人员", "焊工", "电焊工", "检验员", "检测员",
        "试验员", "质检员", "技术员", "工程师", "工艺师", "监督", "监理", "负责人",
        "培训", "考核", "持证", "技能", "经验", "责任人", "人员", "员工", "工人",
    },
    "机": {
        "机器", "设备", "焊机", "电源", "焊枪", "喷嘴", "夹具", "工装", "模具", "刀具", "刃具",
        "仪器", "测量仪器", "计量器具", "试验设备", "测试设备", "定位装置", "定量装置",
        "滚轮", "电极", "钨极", "维护", "保养", "点检", "校准", "精度", "性能",
    },
    "料": {
        "材料", "物料", "原料", "半成品", "成品", "母材", "焊材", "焊丝", "焊条", "焊剂", "焊料",
        "药皮", "填充金属", "熔敷金属", "钢", "铁", "铜", "铝", "镍", "钛", "锌", "镁", "合金",
        "不锈钢", "低碳钢", "高强钢", "保护气", "气体", "氩气", "氦气", "氧气", "氮气", "二氧化碳",
        "成分", "物理性能", "化学性能", "型号", "牌号", "批次", "序列号", "保质期",
    },
    "法": {
        "方法", "工艺", "规程", "规范", "制度", "文件", "工艺文件", "操作规程", "检验指导书", "标准",
        "生产工艺", "设备选择", "首件检验", "控制图", "工艺纪律", "预热", "后热", "坡口", "清理",
        "定位", "安装", "调整", "刃磨", "更换", "操作方法", "管理制度", "验收方法", "验收标准",
    },
    "环": {
        "环境", "温度", "湿度", "照明", "光线", "清洁", "卫生", "现场", "海拔", "污染", "污染度",
        "安全", "环保", "空气", "风速", "水下", "高压", "低温", "高温", "粉尘", "噪声", "5S",
        "定置", "整齐", "有序", "杂物", "职工健康", "生产环境",
    },
    "测": {
        "测量", "检测", "检验", "试验", "测试", "校验", "校准", "确认", "准确度", "精密度", "频次",
        "记录", "测量点", "测量工具", "测试设备", "试验设备", "计量器具", "校准记录", "有效性",
        "电流", "电压", "温度", "压力", "速度", "流量", "长度", "宽度", "高度", "角度", "频率", "尺寸",
    },
}

PERSON_KEEP_TERMS = {
    "申请人", "负责人", "责任人", "操作人", "使用人", "检验人", "审核人", "批准人", "作业人",
    "左手", "右手", "双手", "手指", "手腕", "手臂", "人手", "人员", "员工", "工人", "焊工",
}
PERSON_ROLE_SUFFIXES = ("人", "人员", "者", "员", "工", "师", "长", "组", "班")
PERSON_ROLE_KEYWORDS = CATEGORY_HINTS["人"] | {
    "管理人员", "维修人员", "检查人员", "无损检测人员", "焊接操作人员", "焊接技术人员", "焊接检验人员",
}

CHINESE_SURNAMES = set(
    "赵钱孙李周吴郑王冯陈褚卫蒋沈韩杨朱秦尤许何吕施张孔曹严华金魏陶姜"
    "戚谢邹喻柏水窦章云苏潘葛奚范彭郎鲁韦昌马苗凤花方俞任袁柳酆鲍史唐"
    "费廉岑薛雷贺倪汤滕殷罗毕郝邬安常乐于时傅皮卞齐康伍余元卜顾孟平黄"
    "和穆萧尹姚邵湛汪祁毛禹狄米贝明臧计伏成戴谈宋庞熊纪舒屈项祝董梁杜"
    "阮蓝闵席季麻强贾路娄危江童颜郭梅盛林刁钟徐邱骆高夏蔡田胡凌霍虞万"
    "支柯昝管卢莫经房裘缪干解应宗丁宣邓郁单杭洪包诸左石崔吉龚程嵇邢滑"
    "裴陆荣翁荀羊於惠甄曲家封芮羿储靳汲邴糜松井段富巫乌焦巴弓牧隗山谷"
    "车侯宓蓬全郗班仰秋仲伊宫宁仇栾暴甘斜厉戎祖武符刘景詹束龙叶幸司韶"
    "郜黎蓟薄印宿白怀蒲邰从鄂索咸籍赖卓蔺屠蒙池乔阴鬱胥能苍双闻莘党翟"
    "谭贡劳逄姬申扶堵冉宰郦雍郤璩桑桂濮牛寿通边扈燕冀郏浦尚农温别庄晏"
    "柴瞿阎充慕连茹习宦艾鱼容向古易慎戈廖庾终暨居衡步都耿满弘匡国文寇"
    "广禄阙东殴殳沃利蔚越夔隆师巩厍聂晁勾敖融冷訾辛阚那简饶空曾毋沙乜"
    "养鞠须丰巢关蒯相查后荆红游竺权逯盖益桓公"
)
COMPOUND_SURNAMES = {
    "欧阳", "太史", "端木", "上官", "司马", "东方", "独孤", "南宫", "万俟", "闻人",
    "夏侯", "诸葛", "尉迟", "公羊", "赫连", "澹台", "皇甫", "宗政", "濮阳", "公冶",
    "太叔", "申屠", "公孙", "慕容", "仲孙", "钟离", "长孙", "宇文", "司徒", "鲜于",
    "司空", "闾丘", "子车", "亓官", "司寇", "巫马", "公西", "颛孙", "壤驷", "公良",
    "漆雕", "乐正", "宰父", "谷梁", "拓跋", "夹谷", "轩辕", "令狐", "段干", "百里",
    "呼延", "东郭", "南门", "羊舌", "微生", "左丘", "西门", "第五",
}
NAME_CONTEXT_RE = re.compile(
    r"(作者|审者|编者|主编|副主编|编著|著者|责任编辑|审稿|审核|校对|译者|姓名|签名|编写|参编|"
    r"专家|教授|博士|硕士|院士|参考文献|出版社|\[[JMCDS]\]|［[JMCDS]］)"
)

MATERIAL_KEEP_KEYWORDS = CATEGORY_HINTS["料"]
MATERIAL_GRADE_PATTERNS = [
    r"^H\d+[A-Za-z][A-Za-z0-9-]*$",       # H10MnSi, H0Cr21Ni10
    r"^W\d+[A-Za-z0-9-]*$",              # W707Ni, WC20
    r"^E\d+[A-Za-z0-9-]*$",              # E4303, E5015
    r"^ER\d+[A-Za-z0-9-]*$",             # ER50-6
    r"^Q\d+[A-Za-z0-9-]*$",              # Q235
    r"^[A-Z]{1,3}\d{2,}[A-Za-z0-9-]*$", # 16Mn, 12Cr1MoV, WC20
    r"^\d+[A-Za-z]+[A-Za-z0-9-]*$",      # 316L, 304L, 12Cr1MoV
]

CHEM_FORMULA_DROP = {"CO", "NO", "NO2", "SO2", "SO3"}
CHEM_FORMULA_REVIEW = {"CO2", "O2", "N2", "H2", "Ar", "He", "H2O"}

ELEMENTS = {
    "H", "He", "Li", "Be", "B", "C", "N", "O", "F", "Ne", "Na", "Mg", "Al", "Si", "P", "S", "Cl", "Ar",
    "K", "Ca", "Sc", "Ti", "V", "Cr", "Mn", "Fe", "Co", "Ni", "Cu", "Zn", "Ga", "Ge", "As", "Se", "Br", "Kr",
    "Rb", "Sr", "Y", "Zr", "Nb", "Mo", "Tc", "Ru", "Rh", "Pd", "Ag", "Cd", "In", "Sn", "Sb", "Te", "I", "Xe",
    "Cs", "Ba", "La", "Ce", "Pr", "Nd", "Pm", "Sm", "Eu", "Gd", "Tb", "Dy", "Ho", "Er", "Tm", "Yb", "Lu",
    "Hf", "Ta", "W", "Re", "Os", "Ir", "Pt", "Au", "Hg", "Tl", "Pb", "Bi", "Po", "At", "Rn",
}
SUBSCRIPT_MAP = str.maketrans("₀₁₂₃₄₅₆₇₈₉", "0123456789")

GENERAL_STOP_TERMS = {
    "进行", "要求", "符合", "采用", "确保", "可以", "需要", "相关", "过程", "情况", "使用", "工作",
    "内容", "部分", "方面", "条件", "原因", "问题", "方法", "方式", "影响", "作用", "结果", "规定",
    "图", "表", "公式", "章节", "参考文献", "目录", "附录", "前言", "说明", "数据", "信息",
    "质量", "技术", "特点", "关系", "因素", "措施", "系统", "结构", "控制", "管理",
}
HTML_OCR_NOISE_RE = re.compile(
    r"(rowspan|colspan|</?td|</?tr|</?table|content|image|\.jpg|\.jpeg|\.png|\.pdf|json|base64|http|www\.)",
    re.I,
)


def clean_term(term: Any) -> str:
    term = "" if term is None else str(term)
    term = term.translate(SUBSCRIPT_MAP)
    term = term.strip(" ：:，,。；;、\n\t\r'\"`“”‘’（）()[]【】{}")
    term = re.sub(r"\s+", "", term)
    return term.replace("$", "")


def contains_any(text: str, words: set[str]) -> bool:
    return any(w in text for w in words)


def is_all_cjk(text: str) -> bool:
    return bool(re.fullmatch(r"[\u4e00-\u9fff]+", text))


def has_chinese_surname(text: str) -> bool:
    return bool(text) and (text[:2] in COMPOUND_SURNAMES or text[0] in CHINESE_SURNAMES)


def looks_like_chinese_real_name(term: str, source_text: str = "") -> str | None:
    if term in PERSON_KEEP_TERMS:
        return None
    if term.endswith(PERSON_ROLE_SUFFIXES) or contains_any(term, PERSON_ROLE_KEYWORDS):
        return None
    if not is_all_cjk(term):
        return None

    if 2 <= len(term) <= 4 and has_chinese_surname(term) and NAME_CONTEXT_RE.search(source_text or ""):
        return "疑似真实中文人名：命中姓名/作者/参考文献上下文"
    if 3 <= len(term) <= 4 and term[:2] in COMPOUND_SURNAMES:
        return "疑似真实中文复姓人名"
    # 无上下文的 2~3 字姓名只进入复核，不直接删。
    if 2 <= len(term) <= 3 and has_chinese_surname(term):
        return "REVIEW:无强上下文的疑似中文姓名"
    return None


def looks_like_english_real_name(term: str, source_text: str = "") -> str | None:
    if re.fullmatch(r"[A-Z][a-z]+([·\-\s][A-Z][a-z]+){1,3}", term):
        return "疑似真实英文人名"
    if NAME_CONTEXT_RE.search(source_text or "") and re.fullmatch(r"[A-Za-z][A-Za-z·\-\s.]{2,40}", term):
        return "疑似英文作者/姓名"
    return None


def looks_like_material_grade(term: str) -> bool:
    if any(re.fullmatch(p, term) for p in MATERIAL_GRADE_PATTERNS):
        return True
    if contains_any(term, MATERIAL_KEEP_KEYWORDS) and re.search(r"[A-Za-z0-9]", term):
        return True
    return False


def tokenize_formula(formula: str) -> list[str] | None:
    s = formula.translate(SUBSCRIPT_MAP)
    s = s.replace("（", "(").replace("）", ")")
    s = re.sub(r"[\s·•]", "", s)
    if not re.fullmatch(r"[A-Z][A-Za-z0-9()]*", s):
        return None
    tokens = re.findall(r"[A-Z][a-z]?", s)
    if not tokens:
        return None
    reconstructed = "".join(re.findall(r"[A-Z][a-z]?|\d+|[()]", s))
    if all(t in ELEMENTS for t in tokens) and reconstructed == s:
        return tokens
    return None


def formula_status(term: str, drop_gas_formula: bool = False) -> str | None:
    """返回 DROP / REVIEW / None。"""
    t = clean_term(term)
    if looks_like_material_grade(t):
        return None
    if contains_any(t, MATERIAL_KEEP_KEYWORDS):
        return None
    if t in CHEM_FORMULA_DROP:
        return "DROP"
    if t in CHEM_FORMULA_REVIEW:
        return "DROP" if drop_gas_formula else "REVIEW"
    tokens = tokenize_formula(t)
    if tokens and len(t) <= 10:
        # 氧化物、碳化物等裸化学式不作为材料实体，进入删除。
        return "DROP"
    return None


def looks_like_noise(term: str) -> str | None:
    if len(term) < 2:
        return "过短术语"
    if len(term) > 40:
        return "过长术语"
    if term in GENERAL_STOP_TERMS:
        return "低信息泛词"
    if HTML_OCR_NOISE_RE.search(term):
        return "HTML/OCR/文件残留噪声"
    if re.fullmatch(r"[\d\W_]+", term):
        return "纯数字/纯符号"
    if re.fullmatch(r"第[一二三四五六七八九十百千\d]+[章节篇卷]?", term):
        return "章节号"
    if re.fullmatch(r"[图表]\d+([\-—–]\d+)?", term, flags=re.I):
        return "图表编号"
    if re.fullmatch(r"[A-Za-z]([\-—–][A-Za-z])?区?", term):
        return "单字母区间标记"
    if re.fullmatch(r"[A-Za-z]{1,2}", term):
        return "过短英文缩写"
    return None


def suggest_categories(term: str) -> list[str]:
    hits = []
    for cat, kws in CATEGORY_HINTS.items():
        if contains_any(term, kws):
            hits.append(cat)
    # 常见测量量：即使在焊接工艺中出现，也应提示可归入“测”。
    if re.search(r"(电流|电压|温度|压力|流量|速度|长度|宽度|高度|角度|频率|尺寸|误差|精度)", term):
        if "测" not in hits:
            hits.append("测")
    return hits


def row_text(row: dict[str, Any]) -> str:
    return " ".join(str(row.get(k) or "") for k in ["术语", "归一化术语", "原文片段", "说明", "来源文档", "文本段"])


def decide(row: dict[str, Any], drop_gas_formula: bool = False) -> tuple[str, str]:
    """返回 action, reason。action: KEEP / DROP / REVIEW。"""
    term = clean_term(row.get("术语") or row.get("term"))
    category = clean_term(row.get("类别") or row.get("category"))
    normalized = clean_term(row.get("归一化术语") or row.get("normalized_term"))
    source_text = row_text(row)

    if not term:
        return "DROP", "空术语"
    if category not in ALLOWED_CATEGORIES:
        return "DROP", f"非法类别：{category}"

    noise = looks_like_noise(term)
    if noise:
        return "DROP", noise

    if category == "人":
        reason = looks_like_chinese_real_name(term, source_text)
        if reason:
            return ("REVIEW", reason.replace("REVIEW:", "")) if reason.startswith("REVIEW:") else ("DROP", reason)
        if normalized and normalized != term:
            reason = looks_like_chinese_real_name(normalized, source_text)
            if reason:
                return ("REVIEW", "归一化术语" + reason.replace("REVIEW:", "")) if reason.startswith("REVIEW:") else ("DROP", "归一化术语" + reason)
        reason = looks_like_english_real_name(term, source_text)
        if reason:
            return "DROP", reason

    if category == "料":
        st = formula_status(term, drop_gas_formula=drop_gas_formula)
        if st == "DROP":
            return "DROP", "料类别中的裸化学式/成分符号"
        if st == "REVIEW":
            return "REVIEW", "料类别中的裸气体化学式，需确认是否作为保护气/物料保留"
        if normalized and normalized != term:
            st = formula_status(normalized, drop_gas_formula=drop_gas_formula)
            if st == "DROP":
                return "DROP", "归一化术语为裸化学式/成分符号"
            if st == "REVIEW":
                return "REVIEW", "归一化术语为裸气体化学式，需确认是否作为保护气/物料保留"

    return "KEEP", ""


def rows_to_dicts(ws):
    headers = [clean_term(c.value) for c in ws[1]]
    for idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        values = [cell.value for cell in row]
        yield idx, row, dict(zip(headers, values))


def normalize_key(row_dict: dict[str, Any], dedupe_by_doc_chunk: bool = False) -> tuple:
    term = clean_term(row_dict.get("归一化术语") or row_dict.get("术语"))
    category = clean_term(row_dict.get("类别"))
    if dedupe_by_doc_chunk:
        return (term, category, row_dict.get("来源文档"), str(row_dict.get("文本段")))
    return (term, category)


def style_sheet(ws) -> None:
    fill = PatternFill("solid", fgColor="D9EAF7")
    font = Font(bold=True, color="1F1F1F")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    for col_idx, column in enumerate(ws.columns, start=1):
        max_len = 8
        for cell in column[:300]:
            val = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, min(len(val), 50))
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 8), 50)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def save_table(path: Path, sheet_name: str, headers: list[str], rows: list[list[Any]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]
    ws.append(headers)
    for r in rows:
        ws.append(r)
    style_sheet(ws)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def clean_workbook(
    input_path: Path,
    output_path: Path,
    removed_path: Path,
    review_path: Path,
    category_suspect_path: Path,
    dedupe_by_doc_chunk: bool,
    drop_gas_formula: bool,
) -> None:
    wb_in = load_workbook(input_path, read_only=True, data_only=True)
    ws_in = wb_in.active
    headers = [clean_term(c.value) for c in ws_in[1]]
    if "术语" not in headers or "类别" not in headers:
        raise ValueError(f"未找到必要表头：术语、类别。当前表头为：{headers}")

    kept_rows: list[list[Any]] = []
    removed_rows: list[list[Any]] = []
    review_rows: list[list[Any]] = []
    category_rows: list[list[Any]] = []
    seen: OrderedDict[tuple, int] = OrderedDict()
    dup_count = 0

    for _, src_row, row_dict in rows_to_dicts(ws_in):
        values = [cell.value for cell in src_row]
        action, reason = decide(row_dict, drop_gas_formula=drop_gas_formula)

        if action == "KEEP":
            key = normalize_key(row_dict, dedupe_by_doc_chunk=dedupe_by_doc_chunk)
            if key in seen:
                action, reason = "DROP", "重复术语"
                dup_count += 1

        category = clean_term(row_dict.get("类别"))
        term = clean_term(row_dict.get("术语"))
        hints = [c for c in suggest_categories(term) if c != category]
        if action == "KEEP" and hints:
            category_rows.append(values + ["、".join(hints), f"当前为“{category}”，但术语特征更像：{'、'.join(hints)}"])

        if action == "KEEP":
            kept_rows.append(values)
            seen[normalize_key(row_dict, dedupe_by_doc_chunk=dedupe_by_doc_chunk)] = len(kept_rows)
        elif action == "REVIEW":
            review_rows.append(values + [reason])
        else:
            removed_rows.append(values + [reason])

    save_table(output_path, ws_in.title or "6M术语", headers, kept_rows)
    save_table(removed_path, "removed", headers + ["过滤原因"], removed_rows)
    save_table(review_path, "review", headers + ["复核原因"], review_rows)
    save_table(category_suspect_path, "category_suspect", headers + ["建议类别", "原因"], category_rows)

    print(f"输入：{input_path}")
    print(f"保留：{len(kept_rows)} 条 -> {output_path}")
    print(f"删除：{len(removed_rows)} 条 -> {removed_path}")
    print(f"复核：{len(review_rows)} 条 -> {review_path}")
    print(f"类别疑似错误：{len(category_rows)} 条 -> {category_suspect_path}")
    print(f"其中重复术语：{dup_count} 条")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="筛选/清洗 Welding-6m-Extractor 生成的 6M/5M1E 术语 Excel v2")
    parser.add_argument("input", help="输入 xlsx，例如 1_6m_terms.xlsx")
    parser.add_argument("-o", "--output", default=None, help="清洗后的 xlsx；默认在原文件名后加 .cleaned.v2.xlsx")
    parser.add_argument("--removed", default=None, help="被过滤记录报告 xlsx；默认在原文件名后加 .removed.v2.xlsx")
    parser.add_argument("--review", default=None, help="待人工复核 xlsx；默认在原文件名后加 .review.v2.xlsx")
    parser.add_argument("--category-suspect", default=None, help="类别疑似错误 xlsx；默认在原文件名后加 .category_suspect.v2.xlsx")
    parser.add_argument("--drop-gas-formula", action="store_true", help="删除 CO2、Ar、O2、N2、H2、He 等裸气体化学式，而不是进入复核")
    parser.add_argument("--dedupe-by-doc-chunk", action="store_true", help="去重时保留同一术语在不同来源文档/文本段中的记录")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = Path(args.input)
    output_path = Path(args.output) if args.output else input_path.with_suffix(".cleaned.v2.xlsx")
    removed_path = Path(args.removed) if args.removed else input_path.with_suffix(".removed.v2.xlsx")
    review_path = Path(args.review) if args.review else input_path.with_suffix(".review.v2.xlsx")
    category_suspect_path = Path(args.category_suspect) if args.category_suspect else input_path.with_suffix(".category_suspect.v2.xlsx")
    clean_workbook(
        input_path=input_path,
        output_path=output_path,
        removed_path=removed_path,
        review_path=review_path,
        category_suspect_path=category_suspect_path,
        dedupe_by_doc_chunk=args.dedupe_by_doc_chunk,
        drop_gas_formula=args.drop_gas_formula,
    )


if __name__ == "__main__":
    main()
