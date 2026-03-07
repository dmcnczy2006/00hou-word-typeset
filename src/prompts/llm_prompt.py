"""
LLM 意图解析提示词（统一入口）

所有 LLM 调用的 prompt 由此模块生成，便于维护和修改。

调用关系：
    intent.parser.parse()
        → build_intent_prompt(user_prompt, schema_json, preset_name)
            → rules_loader.format_rules_for_prompt(preset_name)  # 规则片段
            → 返回完整 prompt
"""

from typing import Optional

from ..config.rules_loader import format_rules_for_prompt


# 意图解析的完整 prompt 模板
# 占位符：{user_prompt}、{rules_context}、{schema_json}
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

### 粗体、斜体、下划线
- font_config 中支持 bold、italic、underline（true/false），默认各级标题加粗
- 「正文加粗斜体下划线」→ font_config 中 {{"bold": true, "italic": true, "underline": true}}
- 「所有字都是正的不要斜体」→ 必须添加 scope_rules 中一条 {{"target_scope": "all", "font_config": {{"italic": false}}}}，该条会应用到全文（含所有标题）

### 段落格式
- 行距：line_spacing（如 1.5 表示 1.5 倍行距）
- 段前、段后：space_before、space_after（单位 pt）

{rules_context}

## JSON Schema
{schema_json}

请直接输出 JSON："""


def build_intent_prompt(
    user_prompt: str,
    schema_json: str,
    preset_name: Optional[str] = None,
) -> str:
    """
    构建意图解析的完整 LLM prompt

    调用：intent.parser.parse() 调用本函数
    被调用：format_rules_for_prompt()（rules_loader）

    Args:
        user_prompt: 用户自然语言指令
        schema_json: TypesettingIntent 的 JSON Schema 字符串
        preset_name: 预设名称，用于注入当前预设规则摘要（来自 typesetting_rules.json）

    Returns:
        完整的 prompt 字符串
    """
    rules_context = format_rules_for_prompt(preset_name=preset_name)
    return INTENT_PROMPT_TEMPLATE.format(
        user_prompt=user_prompt.strip(),
        rules_context=rules_context,
        schema_json=schema_json,
    )
