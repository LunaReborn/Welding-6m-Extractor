import re

ALLOWED_CATEGORIES = {"人", "机", "料", "法", "环", "测"}
STOP_WORDS = {
    "进行", "要求", "符合", "采用", "确保", "应", "不得", "必须", "可以", "需要", "相关", "过程", "情况",
    "检查", "确认", "记录", "处理", "使用", "操作", "管理", "工作", "人员", "设备", "材料",
}
NOISE_PATTERNS = [
    re.compile(r"^[\d\W_]+$"),
    re.compile(r"^第[一二三四五六七八九十\d]+$"),
]


def validate_records(records: list[dict]) -> list[dict]:
    valid = []
    seen = set()
    for record in records:
        term = clean_term(str(record.get("term", "")))
        category = str(record.get("category", ""))
        if category not in ALLOWED_CATEGORIES:
            continue
        if not is_valid_term(term):
            continue
        key = (term, category, record.get("document"), record.get("chunk_id"))
        if key in seen:
            continue
        seen.add(key)
        record = dict(record)
        record["term"] = term
        valid.append(record)
    return valid


def clean_term(term: str) -> str:
    term = term.strip(" ：:，,。；;、\n\t\r")
    term = re.sub(r"\s+", "", term)
    return term


def is_valid_term(term: str) -> bool:
    if len(term) < 2 or len(term) > 40:
        return False
    if term in STOP_WORDS:
        return False
    if any(pattern.match(term) for pattern in NOISE_PATTERNS):
        return False
    return True
