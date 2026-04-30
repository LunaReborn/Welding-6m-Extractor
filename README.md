# 焊接文档 6M 专有词汇抽取工具

用于从焊接相关文档中抽取“人、机、料、法、环、测”六类专有词汇，并导出 JSON 和 Excel。

## 目录结构

```text
project/
  cli.py
  engine/
    document_loader.py
    chunker.py
    llm_client.py
    extractor_6m.py
    normalizer.py
    validator.py
    exporter.py
  prompts/
    welding_6m_extract.txt
    welding_6m_review.txt
  outputs/
  data/
```

## 安装

```bash
pip install -r requirements.txt
```

## 使用 Ollama

先启动 Ollama，并拉取模型：

```bash
ollama serve
ollama pull qwen2.5:7b
```

运行：

```bash
python cli.py --file data/氩弧焊标准作业指导书.docx --extractor ollama --model qwen2.5:7b --output outputs
```

启用二阶段复核：

```bash
python cli.py --file data/氩弧焊标准作业指导书.docx --extractor ollama --review --output outputs
```

## 使用 OpenAI API

```bash
set OPENAI_API_KEY=你的key
python cli.py --file data/氩弧焊标准作业指导书.docx --extractor openai --model gpt-4o-mini --output outputs
```

Linux/macOS：

```bash
export OPENAI_API_KEY=你的key
python cli.py --file data/氩弧焊标准作业指导书.docx --extractor openai --model gpt-4o-mini --output outputs
```

## 输出文件

```text
outputs/文档名_6m_terms.json
outputs/文档名_6m_terms.xlsx
```

Excel 字段：

| 字段 | 说明 |
|---|---|
| 术语 | 抽取到的专有词汇 |
| 类别 | 人/机/料/法/环/测 |
| 归一化术语 | 同义词合并后的标准术语 |
| 原文片段 | 术语所在上下文 |
| 来源文档 | 输入文件名 |
| 文本段 | chunk 编号 |
| 置信度 | high/medium/low |
| 说明 | 预留解释字段 |

## 建议

- 如果本地小模型漏抽较多，建议打开 `--review`。
- 如果术语重复较多，在 `engine/normalizer.py` 的 `SYNONYMS` 中持续沉淀同义词。
- 如果抽出普通词较多，在 `engine/validator.py` 的 `STOP_WORDS` 中增加停用词。
