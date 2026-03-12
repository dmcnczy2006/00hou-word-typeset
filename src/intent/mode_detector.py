"""
模式检测器（auto 模式）

根据用户指令，通过 LLM 判断应使用 Y Mode（单次再排版）还是 Z Mode（多步骤工作流）。
"""

import logging
from typing import Literal, Optional

logger = logging.getLogger(__name__)

from ..prompts.workflow_prompt import build_mode_detection_prompt


def detect_mode(
    user_prompt: str,
    llm_connector=None,
) -> Literal["y", "z"]:
    """
    通过 LLM 判断用户指令应使用的模式

    Args:
        user_prompt: 用户自然语言排版指令
        llm_connector: LLM 连接器，若为 None 则使用 OpenAIConnector

    Returns:
        "y" 或 "z"
    """
    if not user_prompt or not user_prompt.strip():
        return "y"

    if llm_connector is None:
        try:
            from llm.connector import OpenAIConnector
            llm = OpenAIConnector()
        except (ImportError, Exception) as e:
            logger.warning("LLM 不可用，模式检测回退到 Y: %s", e)
            return "y"
    else:
        llm = llm_connector

    prompt = build_mode_detection_prompt(user_prompt)
    try:
        response = llm.complete(prompt).strip().lower()
        # 取第一个有效字符
        for c in response:
            if c == "y":
                return "y"
            if c == "z":
                return "z"
        logger.warning("LLM 模式检测响应无法解析: %r，回退到 Y", response[:50])
        return "y"
    except Exception as e:
        logger.warning("LLM 模式检测失败: %s，回退到 Y", e)
        return "y"
