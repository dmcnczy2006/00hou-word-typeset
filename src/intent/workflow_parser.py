"""
工作流解析器（Z Mode）

接收用户自然语言描述的多步骤工作流，调用 LLM 输出结构化 Workflow JSON。

调用关系：
    main.process_document(mode="z")
        → WorkflowParser.parse(user_prompt)
            → prompts.workflow_prompt.build_workflow_prompt()
            → llm.complete(prompt)
            → 解析 JSON 为 Workflow
"""

import json
import logging
import re
from typing import Optional

logger = logging.getLogger(__name__)

from pydantic import ValidationError

from ..schemas.typesetting import Workflow
from ..prompts.workflow_prompt import build_workflow_prompt


class _DefaultMockWorkflowConnector:
    """当 llm 包不可用时的默认 Mock 连接器"""

    def complete(self, prompt: str, schema_hint: Optional[str] = None) -> str:
        return '{"mode": "z", "steps": [{"type": "re_typeset", "preset": "official", "user_override": ""}]}'


class WorkflowParser:
    """
    工作流解析器

    将用户自然语言描述的多步骤工作流解析为 Workflow 结构化配置。
    """

    def __init__(self, llm_connector=None):
        """
        初始化工作流解析器

        Args:
            llm_connector: LLM 连接器实例，若为 None 则使用 OpenAIConnector
        """
        if llm_connector is None:
            try:
                from llm.connector import OpenAIConnector
                self._llm = OpenAIConnector()
            except (ImportError, Exception) as e:
                logger.warning("使用 Mock 连接器（OpenAI 不可用: %s）", e)
                try:
                    from llm.connector import MockLLMConnector
                    self._llm = MockLLMConnector(
                        '{"mode": "z", "steps": [{"type": "re_typeset", "preset": "official", "user_override": ""}]}'
                    )
                except ImportError:
                    self._llm = _DefaultMockWorkflowConnector()
        else:
            self._llm = llm_connector

    def parse(self, user_prompt: str) -> Workflow:
        """
        解析用户描述的工作流为 Workflow

        Args:
            user_prompt: 用户自然语言描述的多步骤工作流

        Returns:
            Workflow 实例
        """
        if not user_prompt or not user_prompt.strip():
            logger.info("用户指令为空，返回默认再排版工作流")
            return Workflow(steps=[{"type": "re_typeset", "preset": "official", "user_override": ""}])

        schema_json = json.dumps(
            Workflow.model_json_schema(),
            ensure_ascii=False,
            indent=2,
        )
        prompt = build_workflow_prompt(
            user_prompt=user_prompt.strip(),
            schema_json=schema_json,
        )

        try:
            response = self._llm.complete(prompt, schema_hint=schema_json)
            logger.debug("LLM 响应长度: %d 字符", len(response))
        except Exception as e:
            logger.warning("LLM 调用失败，回退到默认工作流: %s", e)
            return Workflow(steps=[{"type": "re_typeset", "preset": "official", "user_override": user_prompt}])

        parsed = self._extract_and_validate_json(response)
        if parsed is None:
            logger.warning("LLM 响应解析失败，回退到默认工作流")
            return Workflow(steps=[{"type": "re_typeset", "preset": "official", "user_override": user_prompt}])

        return parsed

    def _extract_and_validate_json(self, response: str) -> Optional[Workflow]:
        """从 LLM 响应中提取 JSON 并校验为 Workflow"""
        text = response.strip()
        if "```json" in text:
            match = re.search(r"```json\s*([\s\S]*?)\s*```", text)
            if match:
                text = match.group(1).strip()
        elif "```" in text:
            match = re.search(r"```\s*([\s\S]*?)\s*```", text)
            if match:
                text = match.group(1).strip()

        try:
            data = json.loads(text)
            return Workflow.model_validate(data)
        except (json.JSONDecodeError, ValidationError):
            return None
