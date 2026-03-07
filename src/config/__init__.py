# 配置模块（接口以 Schema 为准，见 src/schemas/typesetting.py）
from .presets import Config
from .rules_loader import (
    load_rules_document,
    load_rules_json,
    get_preset_names,
    get_preset_rules,
    get_scope_rule,
    get_preset_as_scope_rules,
    get_font_size_reference,
    get_indent_reference,
    format_rules_for_prompt,
)

__all__ = [
    "Config",
    "load_rules_document",
    "load_rules_json",
    "get_preset_names",
    "get_preset_rules",
    "get_scope_rule",
    "get_preset_as_scope_rules",
    "get_font_size_reference",
    "get_indent_reference",
    "format_rules_for_prompt",
]
