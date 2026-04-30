from collections import OrderedDict

SYNONYMS = {
    "TIG焊": "氩弧焊",
    "GTAW": "氩弧焊",
    "钨极氩弧焊": "氩弧焊",
    "氩弧焊设备": "氩弧焊机",
    "TIG焊机": "氩弧焊机",
    "保护气": "氩气",
    "保护气体": "氩气",
}


def normalize_records(records: list[dict]) -> list[dict]:
    merged = OrderedDict()
    for record in records:
        record = dict(record)
        term = record["term"]
        record["normalized_term"] = SYNONYMS.get(term, term)
        key = (record["normalized_term"], record["category"])
        if key not in merged:
            merged[key] = record
        else:
            merged[key]["source_text"] = merge_text(merged[key].get("source_text", ""), record.get("source_text", ""))
            merged[key]["document"] = merge_text(merged[key].get("document", ""), record.get("document", ""), sep="; ")
            merged[key]["chunk_id"] = merge_text(str(merged[key].get("chunk_id", "")), str(record.get("chunk_id", "")), sep=", ")
    return list(merged.values())


def merge_text(a: str, b: str, sep: str = " | ") -> str:
    if not a:
        return b
    if not b or b in a:
        return a
    return f"{a}{sep}{b}"
