"""
预设排版规则配置

提供公文标准、论文标准等常用排版预设，供 IntentParser 和主流程使用。
"""

from typing import Any, Dict, Optional

from ..schemas.typesetting import FontConfig, ParagraphConfig, TypesettingIntent


class Config:
    """
    预设排版规则配置类

    提供标准排版规则预设，支持公文、论文等常见场景。
    """

    # 公文标准（GB/T 9704-2012 党政机关公文格式）
    # 正文：仿宋 GB2312，三号（约 16pt）或小四（12pt）视具体单位要求
    # 此处采用常见的小四仿宋、首行缩进 2 字符
    PRESETS: Dict[str, Dict[str, Any]] = {
        "official": {
            "font": {"name": "仿宋", "size_pt": 12.0, "color": None},
            "paragraph": {
                "first_line_indent": 24.0,  # 2 字符 ≈ 24pt
                "line_spacing": 1.5,
                "space_before": 0.0,
                "space_after": 0.0,
                "alignment": "justify",
            },
            "target_scope": "body",
        },
        # 论文标准（常见学术论文格式，中西文分设）
        "thesis": {
            "font": {
                "name_east_asia": "宋体",
                "name_ascii": "Times New Roman",
                "size_pt": 12.0,
                "color": None,
            },
            "paragraph": {
                "first_line_indent": 24.0,
                "line_spacing": 1.5,
                "space_before": 6.0,
                "space_after": 6.0,
                "alignment": "justify",
            },
            "target_scope": "body",
        },
        # 简洁默认
        "default": {
            "font": {"name": "宋体", "size_pt": 12.0, "color": None},
            "paragraph": {
                "first_line_indent": 24.0,
                "line_spacing": 1.0,
                "space_before": 0.0,
                "space_after": 0.0,
                "alignment": "left",
            },
            "target_scope": "all",
        },
    }

    @classmethod
    def get_preset(cls, name: str) -> Optional[Dict[str, Any]]:
        """
        获取指定名称的预设配置

        Args:
            name: 预设名称，如 "official"、"thesis"、"default"

        Returns:
            预设配置字典，若不存在则返回 None
        """
        return cls.PRESETS.get(name)

    @classmethod
    def preset_to_intent(cls, name: str) -> Optional[TypesettingIntent]:
        """
        将预设配置转换为 TypesettingIntent 对象

        Args:
            name: 预设名称

        Returns:
            TypesettingIntent 实例，若预设不存在则返回 None
        """
        preset = cls.get_preset(name)
        if preset is None:
            return None

        font = preset.get("font")
        paragraph = preset.get("paragraph")

        return TypesettingIntent(
            global_styles=FontConfig(**font) if font else None,
            paragraph_config=ParagraphConfig(**paragraph) if paragraph else None,
            replacements=[],
            target_scope=preset.get("target_scope"),
        )

    @classmethod
    def list_presets(cls) -> list:
        """返回所有可用的预设名称列表"""
        return list(cls.PRESETS.keys())
