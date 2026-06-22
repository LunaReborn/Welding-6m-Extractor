from __future__ import annotations

import json
import os
import socket
import urllib.error
import urllib.request
from dataclasses import dataclass
from typing import Protocol


class LLMClient(Protocol):
    def generate(self, prompt: str) -> str:
        ...


@dataclass
class OllamaClient:
    model: str = "qwen3:8b"
    base_url: str = "https://u778156-903e-6d2946e8.westc.seetacloud.com:8443"
    temperature: float = 0.1
    timeout: int = 180

    def generate(self, prompt: str) -> str:
        url = f"{self.base_url.rstrip('/')}/api/generate"
        payload = {
            "model": self.model,
            "prompt": prompt,
            "stream": False,
            "options": {"temperature": self.temperature},
            "format": "json",
        }
        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"}, method="POST")
        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                body = json.loads(resp.read().decode("utf-8"))
                return body.get("response", "")
        except urllib.error.URLError as exc:
            raise RuntimeError("无法连接 Ollama，请确认 ollama serve 已启动，并且模型已 pull") from exc
        except TimeoutError as exc:
            raise RuntimeError(f"Ollama 请求超时，超过 {self.timeout} 秒未返回") from exc
        except socket.timeout as exc:
            raise RuntimeError(f"Ollama 请求超时，超过 {self.timeout} 秒未返回") from exc


@dataclass
class OpenAIClient:
    model: str = "gpt-4o-mini"
    temperature: float = 0.1
    timeout: int = 180

    def generate(self, prompt: str) -> str:
        try:
            from openai import OpenAI
        except ImportError as exc:
            raise RuntimeError("使用 OpenAI API 需要安装 openai：pip install openai") from exc

        if not os.getenv("OPENAI_API_KEY"):
            raise RuntimeError("缺少 OPENAI_API_KEY 环境变量")

        client = OpenAI(timeout=self.timeout)
        response = client.chat.completions.create(
            model=self.model,
            messages=[{"role": "user", "content": prompt}],
            temperature=self.temperature,
        )
        return response.choices[0].message.content or ""


def build_llm_client(name: str, model: str | None = None, timeout: int = 180) -> LLMClient:
    if name == "ollama":
        return OllamaClient(model=model or "qwen2.5:7b", timeout=timeout)
    if name == "openai":
        return OpenAIClient(model=model or "gpt-4o-mini", timeout=timeout)
    raise ValueError(f"未知 LLM 后端：{name}")
