from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from engine.llm_client import LLMClient

CATEGORIES = ["人", "机", "料", "法", "环", "测"]
PROMPT_DIR = Path(__file__).resolve().parents[1] / "prompts"


def load_prompt(name: str) -> str:
    return (PROMPT_DIR / name).read_text(encoding="utf-8")


def parse_json_object(text: str) -> dict[str, Any]:
    text = text.strip()
    text = re.sub(r"^```(?:json)?", "", text, flags=re.I).strip()
    text = re.sub(r"```$", "", text).strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", text, flags=re.S)
        if not match:
            raise
        return json.loads(match.group(0))


@dataclass
class SixMExtractor:
    client: LLMClient
    review: bool = False

    def extract(self, text: str, document_name: str, chunk_id: int) -> list[dict[str, Any]]:
        prompt_template = load_prompt("welding_6m_extract.txt")
        prompt = prompt_template.replace("{文本段}", text)
        response = self.client.generate(prompt)
        data = parse_json_object(response)
        if self.review:
            data = self._review(text, data)
        return self._to_records(data, text, document_name, chunk_id)

    def _review(self, text: str, data: dict[str, Any]) -> dict[str, Any]:
        review_template = load_prompt("welding_6m_review.txt")
        prompt = review_template.replace("{文本段}", text).replace("{候选结果}", json.dumps(data, ensure_ascii=False, indent=2))
        response = self.client.generate(prompt)
        return parse_json_object(response)

    def _to_records(self, data: dict[str, Any], source_text: str, document_name: str, chunk_id: int) -> list[dict[str, Any]]:
        records: list[dict[str, Any]] = []
        for category in CATEGORIES:
            items = data.get(category, []) or []
            for item in items:
                if isinstance(item, str):
                    term = item.strip()
                    confidence = "high" if term and term in source_text else "medium"
                    source_fragment = _find_fragment(source_text, term)
                    reason = ""
                elif isinstance(item, dict):
                    term = str(item.get("term") or item.get("术语") or "").strip()
                    confidence = str(item.get("confidence") or item.get("置信度") or "medium")
                    source_fragment = str(item.get("source_text") or item.get("原文片段") or _find_fragment(source_text, term))
                    reason = str(item.get("reason") or item.get("理由") or "")
                else:
                    continue
                if not term:
                    continue
                records.append(
                    {
                        "term": term,
                        "category": category,
                        "normalized_term": term,
                        "source_text": source_fragment,
                        "document": document_name,
                        "chunk_id": chunk_id,
                        "confidence": confidence,
                        "reason": reason,
                    }
                )
        return records


def _find_fragment(text: str, term: str, window: int = 40) -> str:
    if not term:
        return ""
    index = text.find(term)
    if index < 0:
        return ""
    start = max(0, index - window)
    end = min(len(text), index + len(term) + window)
    return text[start:end].replace("\n", " ").strip()
