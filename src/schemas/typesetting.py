"""
排版指令 Schema 定义（Schema 优先）

本模块为排版系统的单一数据源，所有配置、规则、意图均由此定义。
- 类型定义：FontConfig、ParagraphConfig、ScopeRule、TypesettingIntent
- 作用范围：TargetScope 枚举
- JSON 兼容：支持 font/paragraph 别名，与 typesetting_rules.json 对齐

后续开发请优先扩展本 Schema，再适配调用方。
"""

from typing import Annotated, List, Literal, Optional, Union

from pydantic import AliasChoices, BaseModel, Field


# ---------------------------------------------------------------------------
# 作用范围（与 typesetting_rules.json scope_mapping 一致）
# ---------------------------------------------------------------------------

TargetScope = Literal[
    "body",
    "heading",
    "heading_1", "heading_2", "heading_3", "heading_4",
    "heading_5", "heading_6", "heading_7", "heading_8", "heading_9",
    "caption",
    "emphasis",
    "all",
]

TARGET_SCOPES: tuple[TargetScope, ...] = (
    "body", "heading",
    "heading_1", "heading_2", "heading_3", "heading_4",
    "heading_5", "heading_6", "heading_7", "heading_8", "heading_9",
    "caption", "emphasis", "all",
)


# ---------------------------------------------------------------------------
# 原子配置（FontConfig、ParagraphConfig）
# ---------------------------------------------------------------------------

class FontConfig(BaseModel):
    """字体配置"""

    name: Optional[str] = Field(
        None,
        description="主字体，未指定 name_east_asia/name_ascii 时同时用于中西文",
    )
    name_east_asia: Optional[str] = Field(
        None,
        description="中文/东亚字体，如：仿宋、宋体、黑体",
    )
    name_ascii: Optional[str] = Field(
        None,
        description="西文字体，如：Times New Roman、Arial",
    )
    size_pt: Optional[float] = Field(None, description="字号（磅），小四号≈12pt")
    color: Optional[str] = Field(None, description="颜色，支持 hex 如 #000000 或颜色名称")
    bold: Optional[bool] = Field(None, description="加粗，默认各级标题为 true")
    italic: Optional[bool] = Field(None, description="斜体")
    underline: Optional[bool] = Field(None, description="下划线")


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


# ---------------------------------------------------------------------------
# 单作用范围规则（ScopeRule）
# 支持 JSON 别名 font/paragraph，与 typesetting_rules.json 对齐
# ---------------------------------------------------------------------------

class ScopeRule(BaseModel):
    """
    单作用范围的排版规则

    用于多 scope 配置，每个 target_scope 一条。
    支持 font/font_config、paragraph/paragraph_config 两种 JSON 键名。
    """

    target_scope: str = Field(
        ...,
        description="作用范围：body、heading_1~9、caption、heading、all",
    )
    font_config: Optional[FontConfig] = Field(
        None,
        validation_alias=AliasChoices("font_config", "font"),
        description="该范围的字体配置",
    )
    paragraph_config: Optional[ParagraphConfig] = Field(
        None,
        validation_alias=AliasChoices("paragraph_config", "paragraph"),
        description="该范围的段落配置",
    )


# ---------------------------------------------------------------------------
# 文本替换规则
# ---------------------------------------------------------------------------

class ReplaceRule(BaseModel):
    """文本替换规则"""

    search: str = Field(..., description="查找内容，支持正则表达式")
    replace: str = Field(..., description="替换内容")
    use_regex: bool = Field(False, description="是否将 search 视为正则表达式")


# ---------------------------------------------------------------------------
# Z Mode 工作流 Schema
# ---------------------------------------------------------------------------

class HeadingChangeStep(BaseModel):
    """标题升降级步骤"""

    type: Literal["heading_change"] = "heading_change"
    from_scope: str = Field(..., description="源 scope，如 heading_2")
    to_scope: str = Field(..., description="目标 scope，如 heading_1")


class ApplyListStep(BaseModel):
    """应用列表步骤"""

    type: Literal["apply_list"] = "apply_list"
    target_scope: str = Field(
        ...,
        description="作用范围：heading、body、all",
    )
    list_style: str = Field(
        ...,
        description="列表样式：multilevel、decimal、bullet",
    )


class ReTypesetStep(BaseModel):
    """再排版步骤"""

    type: Literal["re_typeset"] = "re_typeset"
    preset: Optional[str] = Field("official", description="预设名称")
    user_override: Optional[str] = Field(
        None,
        description="传给 IntentParser 的自然语言，如「正文小四、标题一号」",
    )


class PageMarginsStep(BaseModel):
    """页面边距设置步骤"""

    type: Literal["page_margins"] = "page_margins"
    top_pt: Optional[float] = Field(None, description="上边距（磅），1英寸≈72pt")
    bottom_pt: Optional[float] = Field(None, description="下边距（磅）")
    left_pt: Optional[float] = Field(None, description="左边距（磅）")
    right_pt: Optional[float] = Field(None, description="右边距（磅）")


class HeaderFooterStep(BaseModel):
    """页眉页脚设置步骤。用户不说则无，说得模糊则用默认。odd_even 时奇偶页可分别指定。"""

    type: Literal["header_footer"] = "header_footer"
    header_text: Optional[str] = Field(None, description="页眉文字（odd_even 时作为奇数页默认）")
    footer_text: Optional[str] = Field(None, description="页脚文字（odd_even 时作为奇数页默认）")
    odd_even: bool = Field(False, description="奇偶页不同")
    odd_header_text: Optional[str] = Field(None, description="奇数页页眉，odd_even 时优先于 header_text")
    even_header_text: Optional[str] = Field(None, description="偶数页页眉，odd_even 时使用")
    odd_footer_text: Optional[str] = Field(None, description="奇数页页脚，odd_even 时优先于 footer_text")
    even_footer_text: Optional[str] = Field(None, description="偶数页页脚，odd_even 时使用")
    show_chapter: bool = Field(False, description="显示当前章节（标题）")
    odd_show_chapter: Optional[bool] = Field(None, description="奇数页显示章节，odd_even 时优先")
    even_show_chapter: Optional[bool] = Field(None, description="偶数页显示章节，odd_even 时使用")
    underline: bool = Field(False, description="页眉/页脚下划线")
    odd_underline: Optional[bool] = Field(None, description="奇数页下划线，odd_even 时优先")
    even_underline: Optional[bool] = Field(None, description="偶数页下划线，odd_even 时使用")
    page_number: bool = Field(False, description="添加页码")
    odd_page_number: Optional[bool] = Field(None, description="奇数页页码，odd_even 时优先")
    even_page_number: Optional[bool] = Field(None, description="偶数页页码，odd_even 时使用")
    first_page_different: bool = Field(False, description="首页页眉页脚不同")


class UpdateFieldsStep(BaseModel):
    """更新域步骤"""

    type: Literal["update_fields"] = "update_fields"


class CollapseEmptyLinesStep(BaseModel):
    """将 x 个连续空行改成 y 个空行"""

    type: Literal["collapse_empty_lines"] = "collapse_empty_lines"
    from_count: int = Field(..., ge=2, description="连续空行数量 x（>=2）")
    to_count: int = Field(..., ge=0, description="目标空行数量 y（>=0）")


WorkflowStep = Annotated[
    Union[
        HeadingChangeStep,
        ApplyListStep,
        ReTypesetStep,
        PageMarginsStep,
        HeaderFooterStep,
        UpdateFieldsStep,
        CollapseEmptyLinesStep,
    ],
    Field(discriminator="type"),
]


class Workflow(BaseModel):
    """Z Mode 工作流"""

    mode: Literal["z"] = "z"
    steps: List[WorkflowStep] = Field(
        default_factory=list,
        description="工作流步骤列表，按序执行",
    )


# ---------------------------------------------------------------------------
# 排版意图（TypesettingIntent）
# 支持 scope_rules（主）与 legacy 单 scope 格式（兼容）
# ---------------------------------------------------------------------------

class TypesettingIntent(BaseModel):
    """
    排版意图（LLM 输出 / 预设 / 用户指令解析结果）

    推荐使用 scope_rules；legacy 字段（global_styles + paragraph_config + target_scope）
    用于单 scope 场景或兼容旧格式。调用 to_scope_rules() 可统一为 scope_rules。
    """

    global_styles: Optional[FontConfig] = Field(
        None, description="字体配置（legacy 单 scope 模式）"
    )
    paragraph_config: Optional[ParagraphConfig] = Field(
        None, description="段落配置（legacy 单 scope 模式）"
    )
    target_scope: Optional[str] = Field(
        None,
        description="作用范围（legacy）：body、heading、heading_1~9、caption、all",
    )
    scope_rules: List[ScopeRule] = Field(
        default_factory=list,
        description="多 scope 配置（主模式），优先级高于 legacy 字段",
    )
    replacements: List[ReplaceRule] = Field(
        default_factory=list,
        description="文本替换规则列表",
    )

    def to_scope_rules(self) -> List[ScopeRule]:
        """
        统一为 scope_rules 格式，供 Dispatcher 等调用方使用。

        若 scope_rules 非空则直接返回；否则将 legacy 字段转为单条 ScopeRule。
        """
        if self.scope_rules:
            return self.scope_rules
        scope = self.target_scope or "all"
        if self.global_styles is not None or self.paragraph_config is not None:
            return [
                ScopeRule(
                    target_scope=scope,
                    font_config=self.global_styles,
                    paragraph_config=self.paragraph_config,
                )
            ]
        return []


# ---------------------------------------------------------------------------
# Rules 文档 Schema（对应 typesetting_rules.json）
# ---------------------------------------------------------------------------

class RuleScopeEntry(BaseModel):
    """JSON 中单 scope 规则条目（与 typesetting_rules.json rules.* 结构一致）"""

    model_config = {"extra": "ignore"}

    font: Optional[FontConfig] = Field(None, description="字体配置")
    paragraph: Optional[ParagraphConfig] = Field(None, description="段落配置")


class PresetRules(BaseModel):
    """预设的完整规则集"""

    name: Optional[str] = Field(None, description="预设显示名称")
    description: Optional[str] = Field(None, description="预设说明")
    rules: dict[str, RuleScopeEntry] = Field(
        default_factory=dict,
        description="按 scope 分组的规则，键为 target_scope",
    )

    def to_scope_rules(self) -> List[ScopeRule]:
        """转换为 ScopeRule 列表，供 TypesettingIntent 使用"""
        result = []
        for scope, entry in self.rules.items():
            if scope == "emphasis":
                continue
            if entry.font or entry.paragraph:
                result.append(
                    ScopeRule(
                        target_scope=scope,
                        font_config=entry.font,
                        paragraph_config=entry.paragraph,
                    )
                )
        return result


class RulesDocument(BaseModel):
    """typesetting_rules.json 的完整 Schema"""

    version: Optional[str] = Field(None, description="规则文档版本")
    description: Optional[str] = Field(None, description="文档说明")
    font_size_reference: dict[str, float] = Field(
        default_factory=dict,
        description="字号对照表：中文字号名 -> pt",
    )
    indent_reference: dict[str, float] = Field(
        default_factory=dict,
        description="首行缩进对照：字符数描述 -> pt",
    )
    presets: dict[str, PresetRules] = Field(
        default_factory=dict,
        description="预设名称 -> 预设规则",
    )
    scope_mapping: dict[str, str] = Field(
        default_factory=dict,
        description="scope 标识 -> Word 样式描述",
    )
