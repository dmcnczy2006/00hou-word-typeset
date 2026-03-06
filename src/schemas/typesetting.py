"""
排版指令 Pydantic Schema 定义

用于定义 LLM 输出的结构化排版参数，供 IntentParser 和 CommandDispatcher 使用。
"""

from typing import List, Optional

from pydantic import BaseModel, Field


class FontConfig(BaseModel):
    """字体配置"""

    name: Optional[str] = Field(
        None,
        description="主字体，未指定 name_east_asia/name_ascii 时同时用于中西文",
    )
    name_east_asia: Optional[str] = Field(
        None,
        description="中文/东亚字体，如：仿宋、宋体、黑体。与 name_ascii 配合实现中西文分设",
    )
    name_ascii: Optional[str] = Field(
        None,
        description="西文字体，如：Times New Roman、Arial。与 name_east_asia 配合实现中西文分设",
    )
    size_pt: Optional[float] = Field(None, description="字号（磅），小四号≈12pt")
    color: Optional[str] = Field(None, description="颜色，支持 hex 如 #000000 或颜色名称")


class ParagraphConfig(BaseModel):
    """段落配置"""

    first_line_indent: Optional[float] = Field(
        None, description="首行缩进（磅），2字符≈24pt"
    )
    line_spacing: Optional[float] = Field(None, description="行距倍数，如 1.5 表示 1.5 倍行距")
    space_before: Optional[float] = Field(None, description="段前距（磅）")
    space_after: Optional[float] = Field(None, description="段后距（磅）")
    alignment: Optional[str] = Field(
        None,
        description="对齐方式：left/center/right/justify",
    )


class ReplaceRule(BaseModel):
    """文本替换规则"""

    search: str = Field(..., description="查找内容，支持正则表达式")
    replace: str = Field(..., description="替换内容")
    use_regex: bool = Field(False, description="是否将 search 视为正则表达式")


class TypesettingIntent(BaseModel):
    """
    LLM 输出的完整排版意图

    包含全局样式、段落配置、文本替换规则及作用范围。
    """

    global_styles: Optional[FontConfig] = Field(
        None, description="全局字体样式配置"
    )
    paragraph_config: Optional[ParagraphConfig] = Field(
        None, description="段落格式配置"
    )
    replacements: List[ReplaceRule] = Field(
        default_factory=list,
        description="文本替换规则列表",
    )
    target_scope: Optional[str] = Field(
        None,
        description="作用范围：body=正文/heading=标题/all=全文",
    )
