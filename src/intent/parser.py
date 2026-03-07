"""
意图解析器

接收用户自然语言指令，调用 LLM 输出结构化 TypesettingIntent JSON。
支持从 typesetting_rules.json 注入规则上下文，适配各级标题、正文、强调、题注等。
"""

import json
import logging
import re
from typing import Optional

logger = logging.getLogger(__name__)

from pydantic import ValidationError

from ..schemas.typesetting import FontConfig, ParagraphConfig, ScopeRule, TypesettingIntent
from ..config.presets import Config
from ..config.rules_loader import format_rules_for_prompt

# 智能样式映射：中文字号/字体描述 → 标准参数
FONT_SIZE_MAP = {
    "初号": 42,
    "小初": 36,
    "一号": 26,
    "小一": 24,
    "二号": 22,
    "小二": 18,
    "三号": 16,
    "小三": 15,
    "四号": 14,
    "小四": 12,
    "五号": 10.5,
    "小五": 9,
    "六号": 7.5,
    "小六": 6.5,
    "七号": 5.5,
    "八号": 5,
}

# 首行缩进：字符数 → 磅（1 字符 ≈ 12pt 五号字宽）
def _chars_to_pt(chars: float) -> float:
    return chars * 12.0


# 意图解析的 prompt 模板（rules_context 由 format_rules_for_prompt 动态注入）
INTENT_PROMPT_TEMPLATE = """你是一个 Word 文档排版助手。请根据用户的自然语言指令，输出结构化的 JSON 配置。

## 用户指令
{user_prompt}

## 输出要求
请严格按照以下 JSON Schema 输出，只返回 JSON 对象，不要包含 markdown 代码块或其它说明文字。

### 字体（支持中西文分设）
- 常见中文字体：仿宋、宋体、黑体、楷体、微软雅黑
- 常见西文字体：Times New Roman、Arial
- 若用户要求中西文不同字体，使用 name_east_asia（中文）和 name_ascii（西文），例如：「正文中文仿宋、西文 Times New Roman」→ {{"name_east_asia": "仿宋", "name_ascii": "Times New Roman"}}
- 若仅指定一种字体，使用 name 即可

### 对齐方式
left（左对齐）、center（居中）、right（右对齐）、justify（两端对齐）

{rules_context}

## JSON Schema
{schema_json}

请直接输出 JSON："""


class _DefaultMockConnector:
    """当 llm 包不可用时的默认 Mock 连接器"""

    def complete(self, prompt: str, schema_hint: Optional[str] = None) -> str:
        return '{"global_styles": null, "paragraph_config": null, "replacements": [], "target_scope": "body"}'


class IntentParser:
    """
    意图解析器

    将用户自然语言指令解析为 TypesettingIntent 结构化配置。
    """

    def __init__(self, llm_connector=None):
        """
        初始化意图解析器

        Args:
            llm_connector: LLM 连接器实例，若为 None 则使用 OpenAIConnector（从环境变量读取 API Key）
        """
        if llm_connector is None:
            try:
                from llm.connector import OpenAIConnector
                self._llm = OpenAIConnector()
            except (ImportError, Exception) as e:
                # API Key 未配置或 openai 未安装时回退到 Mock
                logger.warning("使用 Mock 连接器（OpenAI 不可用: %s）", e)
                try:
                    from llm.connector import MockLLMConnector
                    self._llm = MockLLMConnector(
                        '{"global_styles": null, "paragraph_config": null, "replacements": [], "target_scope": "body"}'
                    )
                except ImportError:
                    self._llm = _DefaultMockConnector()
        else:
            self._llm = llm_connector

    def parse(
        self,
        user_prompt: str,
        preset: Optional[str] = None,
        merge_with_preset: bool = True,
    ) -> TypesettingIntent:
        """
        解析用户指令为 TypesettingIntent

        Args:
            user_prompt: 用户自然语言指令
            preset: 预设名称，如 "official"、"thesis"
            merge_with_preset: 是否与预设合并（用户指令覆盖预设）

        Returns:
            TypesettingIntent 实例
        """
        # 1. 获取基础配置（预设或空）
        base_intent = Config.preset_to_intent(preset) if preset else None
        if base_intent is None:
            base_intent = TypesettingIntent()

        # 2. 若用户指令为空，直接返回预设
        if not user_prompt or not user_prompt.strip():
            logger.info("用户指令为空，使用预设配置: %s", preset)
            return base_intent

        # 3. 调用 LLM 解析用户指令
        logger.info("调用 LLM 解析用户指令，预设=%s", preset)
        schema_json = json.dumps(
            TypesettingIntent.model_json_schema(),
            ensure_ascii=False,
            indent=2,
        )
        # 从 typesetting_rules.json 注入规则上下文，使 LLM 适配新 rules
        rules_context = format_rules_for_prompt(preset_name=preset)
        prompt = INTENT_PROMPT_TEMPLATE.format(
            user_prompt=user_prompt.strip(),
            rules_context=rules_context,
            schema_json=schema_json,
        )

        try:
            response = self._llm.complete(prompt, schema_hint=schema_json)
            logger.debug("LLM 响应长度: %d 字符", len(response))
        except Exception as e:
            # LLM 调用失败时回退到预设
            logger.warning("LLM 调用失败，回退到预设: %s", e)
            return base_intent

        # 4. 解析 JSON 并校验
        parsed = self._extract_and_validate_json(response)
        if parsed is None:
            logger.warning("LLM 响应解析失败，回退到预设")
            return base_intent

        # 5. 合并：用户指令覆盖预设
        if merge_with_preset:
            merged = self._merge_intents(base_intent, parsed)
            logger.info("已合并预设与 LLM 解析结果")
            return merged
        return parsed

    def _extract_and_validate_json(self, response: str) -> Optional[TypesettingIntent]:
        """
        从 LLM 响应中提取 JSON 并校验为 TypesettingIntent

        支持从 markdown 代码块中提取。
        """
        text = response.strip()
        # 移除可能的 markdown 代码块
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
            return TypesettingIntent.model_validate(data)
        except (json.JSONDecodeError, ValidationError):
            return None

    def _merge_font_config(
        self,
        base: Optional[FontConfig],
        override: Optional[FontConfig],
    ) -> Optional[FontConfig]:
        """合并字体配置：override 中非空字段覆盖 base"""
        if override is None:
            return base
        if base is None:
            return override
        return FontConfig(
            name=override.name if override.name is not None else base.name,
            name_east_asia=override.name_east_asia if override.name_east_asia is not None else base.name_east_asia,
            name_ascii=override.name_ascii if override.name_ascii is not None else base.name_ascii,
            size_pt=override.size_pt if override.size_pt is not None else base.size_pt,
            color=override.color if override.color is not None else base.color,
        )

    def _merge_paragraph_config(
        self,
        base: Optional[ParagraphConfig],
        override: Optional[ParagraphConfig],
    ) -> Optional[ParagraphConfig]:
        """合并段落配置：override 中非空字段覆盖 base"""
        if override is None:
            return base
        if base is None:
            return override
        return ParagraphConfig(
            first_line_indent=override.first_line_indent if override.first_line_indent is not None else base.first_line_indent,
            line_spacing=override.line_spacing if override.line_spacing is not None else base.line_spacing,
            space_before=override.space_before if override.space_before is not None else base.space_before,
            space_after=override.space_after if override.space_after is not None else base.space_after,
            alignment=override.alignment if override.alignment is not None else base.alignment,
        )

    def _merge_scope_rule(self, base: ScopeRule, override: ScopeRule) -> ScopeRule:
        """合并单条 scope 规则：override 中非空字段覆盖 base"""
        return ScopeRule(
            target_scope=override.target_scope,
            font_config=self._merge_font_config(base.font_config, override.font_config),
            paragraph_config=self._merge_paragraph_config(base.paragraph_config, override.paragraph_config),
        )

    def _merge_intents(
        self,
        base: TypesettingIntent,
        override: TypesettingIntent,
    ) -> TypesettingIntent:
        """
        合并两个意图：预设为底，用户指令覆盖同 scope。

        - 预设的 scope_rules 作为基础
        - 用户提到的 scope 覆盖预设中对应项（同 scope 内 font/paragraph 字段级合并）
        - 用户未提到的 scope 保留预设默认
        """
        # 以预设 scope_rules 为底
        base_map = {sr.target_scope: sr for sr in base.scope_rules}
        # 用户 override 覆盖同 scope
        if override.scope_rules:
            for sr in override.scope_rules:
                if sr.target_scope in base_map:
                    base_map[sr.target_scope] = self._merge_scope_rule(base_map[sr.target_scope], sr)
                else:
                    base_map[sr.target_scope] = sr
        merged_scope_rules = list(base_map.values())

        global_styles = self._merge_font_config(base.global_styles, override.global_styles)
        paragraph_config = override.paragraph_config or base.paragraph_config
        replacements = override.replacements if override.replacements else base.replacements
        target_scope = override.target_scope or base.target_scope

        return TypesettingIntent(
            global_styles=global_styles,
            paragraph_config=paragraph_config,
            replacements=replacements,
            target_scope=target_scope,
            scope_rules=merged_scope_rules,
        )
