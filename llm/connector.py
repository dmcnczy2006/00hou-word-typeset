"""
LLM 连接器抽象

提供统一的 LLM 调用接口，便于切换不同后端（OpenAI、本地 Ollama 等）。
"""

import json
import os
from abc import ABC, abstractmethod
from typing import Optional

try:
    from openai import OpenAI
except ImportError:
    OpenAI = None


class LLMConnector(ABC):
    """LLM 连接器抽象基类"""

    @abstractmethod
    def complete(self, prompt: str, schema_hint: Optional[str] = None) -> str:
        """
        发送 prompt 并获取模型响应

        Args:
            prompt: 完整提示词
            schema_hint: 可选的 JSON Schema 提示，用于约束输出格式

        Returns:
            模型返回的文本
        """
        pass


class OpenAIConnector(LLMConnector):
    """
    OpenAI API 连接器

    使用 OpenAI chat completions API，支持 gpt-4、gpt-3.5-turbo 等模型。
    需设置环境变量 OPENAI_API_KEY。
    """

    def __init__(
        self,
        model: Optional[str] = None,
        api_key: Optional[str] = None,
        base_url: Optional[str] = None,
    ):
        """
        初始化 OpenAI 连接器

        Args:
            model: 模型名称，默认从环境变量 OPENAI_MODEL 读取，未设置时用 deepseek-chat
            api_key: API Key，默认从环境变量 OPENAI_API_KEY 读取
            base_url: API 基础 URL，默认从环境变量 OPENAI_BASE_URL 读取
        """
        if OpenAI is None:
            raise ImportError("请安装 openai: pip install openai")
        self.model = model or os.getenv("OPENAI_MODEL", "deepseek-chat")
        self._client = OpenAI(
            api_key=api_key or os.getenv("OPENAI_API_KEY"),
            base_url=base_url or os.getenv("OPENAI_BASE_URL"),
        )

    def complete(self, prompt: str, schema_hint: Optional[str] = None) -> str:
        """调用 OpenAI API 获取响应"""
        messages = [{"role": "user", "content": prompt}]
        if schema_hint:
            messages.append(
                {
                    "role": "system",
                    "content": f"请严格按照以下 JSON Schema 输出，只返回 JSON，不要包含其他文字。\n{schema_hint}",
                }
            )

        response = self._client.chat.completions.create(
            model=self.model,
            messages=messages,
            temperature=0.2,
        )
        return response.choices[0].message.content or ""


class MockLLMConnector(LLMConnector):
    """
    模拟 LLM 连接器（用于测试）

    返回预设的 JSON，不实际调用 API。
    """

    def __init__(self, mock_response: str = "{}"):
        self.mock_response = mock_response

    def complete(self, prompt: str, schema_hint: Optional[str] = None) -> str:
        return self.mock_response
