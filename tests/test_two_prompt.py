import requests
import json


OLLAMA_URL = "http://localhost:11434/api/generate"
MODEL_NAME = "qwen3:8b"


PROMPT_TEMPLATE = """你是焊接制造领域 5M1E/6M 术语抽取专家。请从文本中抽取“人、机、料、法、环、测”六类术语。

分类标准：
- 人：岗位、角色、责任人、操作/检验/工艺人员。
- 机：生产设备、焊接设备、工装夹具、辅助装置、自动化设备。
- 料：母材、焊材、保护气体、试件、耗材、填充材料。
- 法：焊接方法、工艺规程、工艺参数、操作步骤、焊接顺序。
- 环：环境温度、湿度、风速、通风、洁净度、作业环境、安全环境。
- 测：检测方法、测量项目、检验记录、试验项目、测量/检测工具。

边界规则：
1. 只抽取焊接业务直接相关的专有术语、专业短语或关键短语，不输出整句。
2. category 只能填写以下六个字之一：人、机、料、法、环、测。不要输出“人员”“方法”“环境”等扩展写法。
3. 测量/检测工具优先归“测”，生产/焊接设备优先归“机”。例如：焊缝尺、卡尺、测温仪、探伤仪归“测”；TIG焊机、焊枪、送丝机归“机”。
4. 工艺参数归“法”，例如焊接电流、焊接电压、焊接速度、预热温度、层间温度。
5. 环境条件归“环”，例如环境温度、湿度、风速、通风、洁净度。
6. 气孔、裂纹、夹渣等缺陷词默认不要输出；如果以“气孔检测/裂纹检测”形式作为检测项目出现，可归“测”。
7. 每个术语必须能在原文中找到证据片段。

输出要求：
- 只输出 JSON。
- 不要输出 Markdown。
- 不要输出解释。
- 如果没有术语，输出 {"terms": []}。

输出格式示例：
{
  "terms": [
    {
      "raw_term": "氩弧焊",
      "normalized_term": "TIG焊",
      "category": "法",
      "evidence": "采用氩弧焊进行薄板焊接"
    }
  ]
}

待抽取文本：
{text}
"""


def build_prompt(text: str) -> str:
    return PROMPT_TEMPLATE.replace("{text}", text)


def call_ollama(prompt: str) -> str:
    payload = {
        "model": MODEL_NAME,
        "prompt": prompt,
        "stream": False,
        "format": "json",
        "options": {
            "temperature": 0.1,
            "top_p": 0.9
        }
    }

    response = requests.post(
        OLLAMA_URL,
        json=payload,
        timeout=180
    )

    if response.status_code != 200:
        raise RuntimeError(
            f"Ollama 请求失败，状态码：{response.status_code}\n"
            f"返回内容：{response.text}"
        )

    return response.json().get("response", "")


def extract_terms(text: str) -> dict:
    prompt = build_prompt(text)
    content = call_ollama(prompt)

    try:
        result = json.loads(content)
    except json.JSONDecodeError:
        return {
            "terms": [],
            "raw_response": content,
            "error": "模型输出不是合法 JSON"
        }

    if not isinstance(result, dict):
        return {
            "terms": [],
            "raw_response": content,
            "error": "模型输出不是 JSON 对象"
        }

    if "terms" not in result:
        result["terms"] = []

    return result


if __name__ == "__main__":
    text = """
    焊工必须持证上岗，作业前检查TIG焊机、焊枪、气瓶和流量计。
    母材为304不锈钢板，焊材采用ER308L焊丝，保护气体为99.99%氩气。
    焊接前应清理坡口油污和氧化皮，采用氩弧焊进行薄板焊接。
    焊接电流控制在80A到100A，层间温度不超过150℃。
    作业现场应保持良好通风，环境湿度不宜过高。
    焊后使用焊缝尺进行外观检查，并填写检验记录。
    """

    result = extract_terms(text)

    print(json.dumps(result, ensure_ascii=False, indent=2))