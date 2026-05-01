import argparse
import json
import logging
import time
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
    parser.add_argument("--save-every", type=int, default=50, help="每处理多少段保存一次；设为 0 表示只在结束时保存")
    parser.add_argument("--max-retries", type=int, default=1, help="请求异常时额外重试次数；默认 1 表示失败后再请求一次")
    parser.add_argument("--retry-sleep", type=float, default=2.0, help="请求异常重试前等待秒数")
    parser.add_argument("--timeout", type=int, default=180, help="单次 LLM 请求超时时间，单位秒")
    parser.add_argument("--log-level", default="INFO", choices=["DEBUG", "INFO", "WARNING", "ERROR"])
    return parser.parse_args()


def save_outputs(records: list[dict], json_path: Path, xlsx_path: Path, processed_chunks: int | None = None) -> list[dict]:
    """校验、归一化并保存当前累计结果。返回最终保存的记录。"""
    valid_records = validate_records(records)
    final_records = normalize_records(valid_records)
    export_json(final_records, json_path)
    export_xlsx(final_records, xlsx_path)
    if processed_chunks is None:
        logging.info("保存完成：共输出 %s 条术语", len(final_records))
    else:
        logging.info("阶段保存：已处理 %s 段，当前输出 %s 条术语", processed_chunks, len(final_records))
    return final_records


def extract_one_chunk(extractor: SixMExtractor, chunk: str, document_name: str, chunk_id: int, max_retries: int, retry_sleep: float) -> list[dict]:
    """抽取单个文本段。

    - 请求类异常：按 max_retries 重试，默认失败后再请求一次。
    - JSON 解析失败：直接跳过当前段，不重试，避免卡住长任务。
    """
    attempts = max(0, max_retries) + 1
    for attempt in range(1, attempts + 1):
        try:
            return extractor.extract(text=chunk, document_name=document_name, chunk_id=chunk_id)
        except json.JSONDecodeError as exc:
            logging.warning("第 %s 段 JSON 解析失败，已跳过：%s", chunk_id, exc)
            return []
        except Exception as exc:
            if attempt < attempts:
                logging.warning(
                    "第 %s 段请求异常，准备重试 %s/%s：%s",
                    chunk_id,
                    attempt,
                    max_retries,
                    exc,
                )
                if retry_sleep > 0:
                    time.sleep(retry_sleep)
                continue
            logging.error("第 %s 段请求连续失败，已跳过：%s", chunk_id, exc)
            return []
    return []


def main():
    args = parse_args()
    logging.basicConfig(level=getattr(logging, args.log_level), format="%(asctime)s - %(levelname)s - %(message)s")

    input_path = Path(args.file)
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    text = load_document(input_path)
    chunks = chunk_text(text, max_chars=args.chunk_size, overlap=args.chunk_overlap)
    logging.info("文档解析完成：%s，共 %s 个文本段", input_path.name, len(chunks))

    client = build_llm_client(args.extractor, model=args.model, timeout=args.timeout)
    extractor = SixMExtractor(client=client, review=args.review)

    stem = input_path.stem
    json_path = output_dir / f"{stem}_6m_terms.json"
    xlsx_path = output_dir / f"{stem}_6m_terms.xlsx"

    raw_records: list[dict] = []
    for idx, chunk in enumerate(chunks, start=1):
        logging.info("抽取进度：%s/%s", idx, len(chunks))
        records = extract_one_chunk(
            extractor=extractor,
            chunk=chunk,
            document_name=input_path.name,
            chunk_id=idx,
            max_retries=args.max_retries,
            retry_sleep=args.retry_sleep,
        )
        raw_records.extend(records)

        if args.save_every > 0 and idx % args.save_every == 0:
            save_outputs(raw_records, json_path, xlsx_path, processed_chunks=idx)

    final_records = save_outputs(raw_records, json_path, xlsx_path)

    logging.info("完成：共输出 %s 条术语", len(final_records))
    logging.info("JSON：%s", json_path)
    logging.info("Excel：%s", xlsx_path)


if __name__ == "__main__":
    main()
