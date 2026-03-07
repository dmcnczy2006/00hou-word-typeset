"""
排版规则加载器

从 typesetting_rules.json 加载规则，统一返回 Schema 类型（RulesDocument、ScopeRule 等）。
接口以 Schema 为准，与 src/schemas/typesetting.py 对齐。
"""

import json
import logging
from pathlib import Path
from typing import List, Optional

from ..schemas.typesetting import (
    PresetRules,
    RulesDocument,
    ScopeRule,
)

logger = logging.getLogger(__name__)

# 默认规则文件路径（相对于项目根目录）
DEFAULT_RULES_PATH = Path(__file__).resolve().parent.parent.parent / "typesetting_rules.json"


def load_rules_document(path: Optional[Path] = None) -> RulesDocument:
    """
    加载排版规则文档，返回 Schema 类型。

    Args:
        path: JSON 文件路径，若为 None 则使用默认路径

    Returns:
        RulesDocument 实例，加载失败时返回空文档
    """
    p = path or DEFAULT_RULES_PATH
    if not p.exists():
        logger.warning("排版规则文件不存在: %s", p)
        return RulesDocument()

    try:
        with open(p, encoding="utf-8") as f:
            data = json.load(f)
        logger.debug("已加载排版规则: %s", p)
        return RulesDocument.model_validate(data)
    except (json.JSONDecodeError, Exception) as e:
        logger.warning("加载排版规则失败: %s", e)
        return RulesDocument()


def get_preset_names(rules: Optional[RulesDocument] = None) -> List[str]:
    """
    获取所有预设名称列表。

    Args:
        rules: 规则文档，若为 None 则自动加载

    Returns:
        预设名称列表
    """
    doc = rules or load_rules_document()
    return list(doc.presets.keys())


def get_preset_rules(
    preset_name: str,
    rules: Optional[RulesDocument] = None,
) -> Optional[PresetRules]:
    """
    获取指定预设的完整规则（Schema 类型）。

    Args:
        preset_name: 预设名称，如 "official"、"thesis"、"default"
        rules: 规则文档，若为 None 则自动加载

    Returns:
        PresetRules 实例，若预设不存在则返回 None
    """
    doc = rules or load_rules_document()
    return doc.presets.get(preset_name)


def get_scope_rule(
    preset_name: str,
    scope: str,
    rules: Optional[RulesDocument] = None,
) -> Optional[ScopeRule]:
    """
    获取指定预设、指定作用范围的单条规则（Schema 类型）。

    Args:
        preset_name: 预设名称
        scope: 作用范围，如 "heading_1"、"body"、"caption"
        rules: 规则文档

    Returns:
        ScopeRule 实例，若不存在则返回 None
    """
    preset = get_preset_rules(preset_name, rules)
    if preset is None:
        return None
    entry = preset.rules.get(scope)
    if entry is None or (entry.font is None and entry.paragraph is None):
        return None
    return ScopeRule(
        target_scope=scope,
        font_config=entry.font,
        paragraph_config=entry.paragraph,
    )


def get_preset_as_scope_rules(
    preset_name: str,
    rules: Optional[RulesDocument] = None,
) -> List[ScopeRule]:
    """
    将预设规则转换为 ScopeRule 列表（Schema 类型）。

    Args:
        preset_name: 预设名称
        rules: 规则文档

    Returns:
        ScopeRule 列表，不含 emphasis
    """
    preset = get_preset_rules(preset_name, rules)
    if preset is None:
        return []
    return preset.to_scope_rules()


def format_rules_for_prompt(
    preset_name: Optional[str] = None,
    rules: Optional[RulesDocument] = None,
) -> str:
    """
    将预设规则格式化为 LLM prompt 可读的文本。

    Args:
        preset_name: 预设名称，若为 None 则返回通用说明
        rules: 规则文档

    Returns:
        格式化的规则文本，用于注入 LLM 提示词
    """
    doc = rules or load_rules_document()
    lines = []

    # 作用范围说明
    if doc.scope_mapping:
        lines.append("### 作用范围（target_scope）")
        for scope, desc in doc.scope_mapping.items():
            lines.append(f"- {scope}: {desc}")
        lines.append("- heading: 全部标题（兼容旧格式）")
        lines.append("- all: 全文")
        lines.append("")

    # 字号与缩进参考
    if doc.font_size_reference:
        lines.append("### 字号映射（size_pt）")
        for name, pt in list(doc.font_size_reference.items())[:12]:
            lines.append(f"- {name} = {pt}pt")
    if doc.indent_reference:
        lines.append("")
        lines.append("### 首行缩进（first_line_indent，单位 pt）")
        for name, pt in doc.indent_reference.items():
            lines.append(f"- {name} ≈ {pt}pt")
    lines.append("")

    # 多 scope 输出说明
    lines.append("### 多 scope 输出（重要）")
    lines.append("当用户指定多个元素（如「正文...；一级标题...」）时，必须使用 scope_rules 数组，每个元素一条配置。")
    lines.append('示例：scope_rules: [{"target_scope": "body", "font_config": {"name": "仿宋", "size_pt": 12}, "paragraph_config": {"first_line_indent": 24, "line_spacing": 1.5}}, {"target_scope": "heading_1", "font_config": {"name": "黑体", "size_pt": 22}, "paragraph_config": {"alignment": "center"}}]')
    lines.append("单 scope 时仍可使用 global_styles + paragraph_config + target_scope。")
    lines.append("")

    # 当前预设的规则摘要
    if preset_name:
        preset = get_preset_rules(preset_name, doc)
        if preset and preset.rules:
            lines.append(f"### 当前预设「{preset_name}」规则摘要")
            for scope, entry in preset.rules.items():
                if entry.font or entry.paragraph:
                    parts = []
                    if entry.font:
                        name = entry.font.name or entry.font.name_east_asia or "?"
                        size = entry.font.size_pt
                        parts.append(f"{name} {size}pt" if size else name)
                    if entry.paragraph:
                        if entry.paragraph.alignment:
                            parts.append(f"对齐:{entry.paragraph.alignment}")
                        if entry.paragraph.first_line_indent:
                            parts.append(f"缩进:{entry.paragraph.first_line_indent}pt")
                    lines.append(f"- {scope}: {' | '.join(parts)}")
            lines.append("")

    return "\n".join(lines)


def get_font_size_reference(rules: Optional[RulesDocument] = None) -> dict:
    """获取字号对照表（中文字号名 -> pt）"""
    doc = rules or load_rules_document()
    return doc.font_size_reference


def get_indent_reference(rules: Optional[RulesDocument] = None) -> dict:
    """获取首行缩进对照表（字符数描述 -> pt）"""
    doc = rules or load_rules_document()
    return doc.indent_reference


# 兼容旧接口：load_rules_json 返回 dict（供需要原始 dict 的调用方）
def load_rules_json(path: Optional[Path] = None) -> dict:
    """
    加载排版规则 JSON 为原始 dict（兼容接口）。

    推荐使用 load_rules_document() 获取 Schema 类型。
    """
    doc = load_rules_document(path)
    return doc.model_dump(exclude_none=True)
