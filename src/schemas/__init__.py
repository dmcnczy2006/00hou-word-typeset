# 排版指令 Schema 模块（Schema 优先，后续开发请优先扩展 typesetting.py）
from .typesetting import (
    FontConfig,
    ParagraphConfig,
    PresetRules,
    ReplaceRule,
    RuleScopeEntry,
    RulesDocument,
    ScopeRule,
    TargetScope,
    TARGET_SCOPES,
    TypesettingIntent,
)

__all__ = [
    "FontConfig",
    "ParagraphConfig",
    "PresetRules",
    "ReplaceRule",
    "RuleScopeEntry",
    "RulesDocument",
    "ScopeRule",
    "TargetScope",
    "TARGET_SCOPES",
    "TypesettingIntent",
]
