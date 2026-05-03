from ollama import Client
import json

# -------------------------- 配置区 --------------------------
# 替换为你的Ollama服务URL
OLLAMA_SERVICE_URL = "https://127.0.0.1:11434"
MODEL_NAME = "qwen3.5:8=27b"

# 待抽取的焊接文档
welding_document = """
1. 焊接作业人员要求：所有参与本工程焊接的焊工，必须持有特种设备焊接作业人员证，经焊接工艺评定合格后方可上岗，无损检测人员需具备RT/UT二级及以上资质。
2. 焊接设备与工装：采用ZX7-400逆变直流焊机、WSE-315交直流氩弧焊机，配套TIG焊枪、MIG焊枪，焊接变位机、坡口机，使用焊缝检验尺、里氏硬度计进行检测。
3. 焊接材料：母材为Q355B低合金高强度结构钢，焊材选用ER50-6实心焊丝、E5015碱性焊条，保护气体为80%Ar+20%CO2混合气体，焊剂为HJ431熔炼型焊剂。
"""
# -----------------------------------------------------------

# 初始化客户端，指定远程URL
client = Client(host=OLLAMA_SERVICE_URL)

# 核心抽取调用
response = client.chat(
    model=MODEL_NAME,
    messages=[
        {
            "role": "system",
            "content": """
            你是专业的焊接工程文档术语抽取助手，严格遵守以下规则：
            1. 仅从用户提供的焊接文档中，提取焊接领域专有术语，禁止提取通用词汇；
            2. 严格按照「人、机、料、法、环」5个维度分类；
            3. 严格输出标准JSON格式，无任何额外解释文本，结构固定为：
            {"人": [], "机": [], "料": [], "法": [], "环": []}
            """
        },
        {"role": "user", "content": welding_document}
    ],
    format="json",
    stream=False,
    options={"temperature": 0.1, "top_p": 0.3}
)

# 解析结果
try:
    extract_result = json.loads(response["message"]["content"])
    print("===== 抽取结果 =====")
    print(json.dumps(extract_result, ensure_ascii=False, indent=2))
except json.JSONDecodeError:
    print("解析失败，模型输出：")
    print(response["message"]["content"])