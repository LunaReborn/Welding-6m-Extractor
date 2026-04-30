import requests
import json


OLLAMA_URL = "http://localhost:11434/api/generate"
MODEL_NAME = "qwen3:8b"


def extract_6m_terms(text: str) -> dict:
    prompt = f"""
你是一个焊接工艺文档信息抽取专家。

请从下面的焊接文档文本中抽取“人、机、料、法、环、测”六类专有词汇。

分类说明：
1. 人：人员、岗位、资质、操作角色
2. 机：设备、工具、仪器、工装
3. 料：材料、焊材、气体、零件、母材
4. 法：工艺方法、操作步骤、工艺参数、标准规范
5. 环：环境条件、作业环境、安全环境
6. 测：检测方法、检测工具、检验项目、质量要求

要求：
- 只抽取原文中明确出现的词汇
- 不要编造
- 不要解释
- 输出严格 JSON 格式
- 每一类结果使用列表
- 没有内容则返回空列表

文本如下：
{text}

请按照以下 JSON 格式输出：
{{
  "人": [],
  "机": [],
  "料": [],
  "法": [],
  "环": [],
  "测": []
}}
"""

    payload = {
        "model": MODEL_NAME,
        "prompt": prompt,
        "stream": False,
        "format": "json"
    }

    response = requests.post(
        OLLAMA_URL,
        json=payload,
        timeout=180
    )
    response.raise_for_status()

    result = response.json()
    content = result.get("response", "")

    try:
        return json.loads(content)
    except json.JSONDecodeError:
        return {
            "raw_response": content,
            "error": "模型输出不是合法 JSON"
        }


if __name__ == "__main__":
    text = """
    焊工必须持证上岗，作业前检查氩弧焊机、焊枪、气瓶和流量计。
    焊接材料采用ER308L焊丝，保护气体为99.99%氩气。
    焊接前应清理坡口油污和氧化皮，采用TIG焊工艺。
    作业现场应保持良好通风，相对湿度不宜过高。
    焊后进行外观检查和焊缝尺寸检测。
    """

    result = extract_6m_terms(text)

    print(json.dumps(result, ensure_ascii=False, indent=2))