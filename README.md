# Welding 6M Extractor

从焊接相关文档中抽取“人、机、料、法、环、测”六类专有词汇，支持 Ollama 和 OpenAI API。

## 安装

```bash
pip install -r requirements.txt
```

## 基本运行

```bash
python cli.py --file data/sample_welding.txt --extractor ollama --model qwen2.5:7b --output outputs
```

## 推荐运行方式

长文档建议指定“每多少段保存一次”，避免中途异常导致结果全部丢失：

```bash
python cli.py \
  --file data/sample_welding.txt \
  --extractor ollama \
  --model qwen2.5:7b \
  --output outputs \
  --save-every 50 \
  --max-retries 1 \
  --timeout 180
```

含义：

- `--save-every 50`：每处理 50 个文本段保存一次 JSON 和 Excel。
- `--max-retries 1`：请求异常时额外重试 1 次，也就是最多请求 2 次。
- `--timeout 180`：单次 LLM 请求最多等待 180 秒。
- `--save-every 0`：只在全部结束时保存。

## 异常处理策略

当前版本适合长文档批处理：

1. 请求异常，例如超时、网络波动、Ollama 偶发无响应：默认重试一次。
2. 重试后仍失败：跳过当前文本段，继续下一段。
3. JSON 解析失败：不重试，直接跳过当前文本段，继续下一段。
4. 按 `--save-every` 指定的段数增量保存，最后再保存一次完整结果。

## 输出

```text
outputs/
  sample_welding_6m_terms.json
  sample_welding_6m_terms.xlsx
```

Excel 字段：

- 术语
- 类别
- 归一化术语
- 原文片段
- 来源文档
- 文本段
- 置信度
- 说明
