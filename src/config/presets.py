"""
预设排版规则配置

从 typesetting_rules.json 加载预设，支持公文、论文等场景。
用户未提及的 scope 将按预设默认规则排版。
"""

from typing import Optional

from ..schemas.typesetting import FontConfig, ParagraphConfig, ScopeRule, TypesettingIntent
from .rules_loader import get_preset_as_scope_rules, get_preset_names


class Config:
    """
    预设排版规则配置类

    优先从 typesetting_rules.json 加载；若预设不存在或文件缺失，回退到内置简版。
    """

    # 内置简版（typesetting_rules.json 缺失或预设不存在时回退）
    _FALLBACK: dict[str, TypesettingIntent] = {}

    @classmethod
    def _build_fallback(cls) -> None:
        """按需构建内置回退预设"""
        if cls._FALLBACK:
            return
        cls._FALLBACK = {
            "official": TypesettingIntent(
                global_styles=FontConfig(name="仿宋", size_pt=12.0),
                paragraph_config=ParagraphConfig(
                    first_line_indent=24.0,
                    line_spacing=1.5,
                    alignment="justify",
                ),
                target_scope="body",
            ),
            "thesis": TypesettingIntent(
                global_styles=FontConfig(
                    name_east_asia="宋体",
                    name_ascii="Times New Roman",
                    size_pt=12.0,
                ),
                paragraph_config=ParagraphConfig(
                    first_line_indent=24.0,
                    line_spacing=1.5,
                    space_before=6.0,
                    space_after=6.0,
                    alignment="justify",
                ),
                target_scope="body",
            ),
            "default": TypesettingIntent(
                global_styles=FontConfig(name="宋体", size_pt=12.0),
                paragraph_config=ParagraphConfig(
                    first_line_indent=24.0,
                    line_spacing=1.0,
                    alignment="left",
                ),
                target_scope="all",
            ),
        }

    @classmethod
    def preset_to_intent(cls, name: str) -> Optional[TypesettingIntent]:
        """
        将预设转换为 TypesettingIntent

        优先从 typesetting_rules.json 加载完整 scope_rules（各级标题、正文、题注等）；
        若不存在则回退到内置简版（仅 body/all）。

        Args:
            name: 预设名称，如 "official"、"thesis"、"default"

        Returns:
            TypesettingIntent 实例，若预设不存在则返回 None
        """
        scope_rules = get_preset_as_scope_rules(name)
        if scope_rules:
            return TypesettingIntent(scope_rules=scope_rules, replacements=[])

        cls._build_fallback()
        return cls._FALLBACK.get(name)

    @classmethod
    def list_presets(cls) -> list[str]:
        """返回所有可用的预设名称列表"""
        names = get_preset_names()
        if names:
            return names
        cls._build_fallback()
        return list(cls._FALLBACK.keys())
