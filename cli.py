import argparse
import logging
from pathlib import Path

from engine.chunker import chunk_text
from engine.document_loader import load_document
from engine.extractor_6m import SixMExtractor
from engine.exporter import export_json, export_xlsx
from engine.llm_client import build_llm_client
from engine.normalizer import normalize_records
from engine.validator import validate_records


def parse_args():
    parser = argparse.ArgumentParser(description="焊接文档 6M 专有词汇抽取工具")
    parser.add_argument("--file", required=True, help="输入文档路径，支持 .txt/.md/.docx/.pdf")
    parser.add_argument("--extractor", default="ollama", choices=["ollama", "openai"], help="LLM 后端")
    parser.add_argument("--model", default=None, help="模型名称；ollama 默认 qwen2.5:7b，openai 默认 gpt-4o-mini")
    parser.add_argument("--output", default="outputs", help="输出目录")
    parser.add_argument("--chunk-size", type=int, default=1500, help="每段最大字符数")
    parser.add_argument("--chunk-overlap", type=int, default=120, help="分段重叠字符数")
    parser.add_argument("--review", action="store_true", help="启用二阶段复核，速度较慢但更稳")
    parser.add_argument("--log-level", default="INFO", choices=["DEBUG", "INFO", "WARNING", "ERROR"])
    return parser.parse_args()


def main():
    args = parse_args()
    logging.basicConfig(level=getattr(logging, args.log_level), format="%(asctime)s - %(levelname)s - %(message)s")

    input_path = Path(args.file)
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    text = load_document(input_path)
    chunks = chunk_text(text, max_chars=args.chunk_size, overlap=args.chunk_overlap)
    logging.info("文档解析完成：%s，共 %s 个文本段", input_path.name, len(chunks))

    client = build_llm_client(args.extractor, model=args.model)
    extractor = SixMExtractor(client=client, review=args.review)

    raw_records = []
    for idx, chunk in enumerate(chunks, start=1):
        logging.info("抽取进度：%s/%s", idx, len(chunks))
        raw_records.extend(
            extractor.extract(
                text=chunk,
                document_name=input_path.name,
                chunk_id=idx,
            )
        )

    valid_records = validate_records(raw_records)
    final_records = normalize_records(valid_records)

    stem = input_path.stem
    json_path = output_dir / f"{stem}_6m_terms.json"
    xlsx_path = output_dir / f"{stem}_6m_terms.xlsx"
    export_json(final_records, json_path)
    export_xlsx(final_records, xlsx_path)

    logging.info("完成：共输出 %s 条术语", len(final_records))
    logging.info("JSON：%s", json_path)
    logging.info("Excel：%s", xlsx_path)


if __name__ == "__main__":
    main()
