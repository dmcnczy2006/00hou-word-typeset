"""
Word 排版引擎

封装 python-docx 操作，提供字体、段落、文本替换等排版原子能力。
"""

import logging
import re
from typing import Optional

logger = logging.getLogger(__name__)

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt

from ..schemas.typesetting import FontConfig, ParagraphConfig


# 中文字体名到 Word 内部名称的映射（智能样式映射）
FONT_NAME_MAP = {
    "仿宋": "仿宋",
    "仿宋_GB2312": "仿宋_GB2312",
    "宋体": "宋体",
    "黑体": "黑体",
    "楷体": "楷体",
    "微软雅黑": "微软雅黑",
    "Times New Roman": "Times New Roman",
    "Arial": "Arial",
}


def _parse_color(color_str: Optional[str]):
    """
    解析颜色字符串为 RGBColor

    支持 #RRGGBB 或 #RGB 格式，以及常见颜色名称。
    """
    if not color_str:
        return None
    from docx.shared import RGBColor

    color_str = color_str.strip()
    # Hex 格式
    if color_str.startswith("#"):
        hex_val = color_str[1:]
        if len(hex_val) == 6:
            r = int(hex_val[0:2], 16)
            g = int(hex_val[2:4], 16)
            b = int(hex_val[4:6], 16)
            return RGBColor(r, g, b)
        elif len(hex_val) == 3:
            r = int(hex_val[0] * 2, 16)
            g = int(hex_val[1] * 2, 16)
            b = int(hex_val[2] * 2, 16)
            return RGBColor(r, g, b)
    # 常见颜色名称
    color_names = {
        "black": (0, 0, 0),
        "white": (255, 255, 255),
        "red": (255, 0, 0),
        "green": (0, 128, 0),
        "blue": (0, 0, 255),
    }
    lower = color_str.lower()
    if lower in color_names:
        r, g, b = color_names[lower]
        return RGBColor(r, g, b)
    return None


def _alignment_from_str(align: Optional[str]):
    """将字符串对齐方式转换为 WD_ALIGN_PARAGRAPH 枚举"""
    if not align:
        return None
    mapping = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    return mapping.get(align.lower())


# target_scope -> Word 内置样式名（python-docx 使用英文名）
SCOPE_TO_STYLE_NAMES: dict[str, list[str]] = {
    "body": ["Normal"],
    "heading": ["Heading 1", "Heading 2", "Heading 3", "Heading 4", "Heading 5", "Heading 6", "Heading 7", "Heading 8", "Heading 9"],
    "heading_1": ["Heading 1"],
    "heading_2": ["Heading 2"],
    "heading_3": ["Heading 3"],
    "heading_4": ["Heading 4"],
    "heading_5": ["Heading 5"],
    "heading_6": ["Heading 6"],
    "heading_7": ["Heading 7"],
    "heading_8": ["Heading 8"],
    "heading_9": ["Heading 9"],
    "caption": ["Caption"],
    "all": ["Normal", "Heading 1", "Heading 2", "Heading 3", "Heading 4", "Heading 5", "Heading 6", "Heading 7", "Heading 8", "Heading 9", "Caption"],
}


def _get_style_names_for_scope(target_scope: str) -> list[str]:
    """根据 target_scope 返回要修改的 Word 样式名列表"""
    return SCOPE_TO_STYLE_NAMES.get(target_scope, ["Normal"])


def _resolve_font(raw: Optional[str]) -> Optional[str]:
    """解析字体名（智能映射）"""
    if not raw:
        return None
    return FONT_NAME_MAP.get(raw, raw)


# 主题字体属性：在 Heading 等样式中会覆盖显式字体，需移除后显式字体才生效
# 注：OOXML 中为 cstheme（小写）
_THEME_FONT_ATTRS = (
    qn("w:asciiTheme"),
    qn("w:eastAsiaTheme"),
    qn("w:hAnsiTheme"),
    qn("w:cstheme"),
)


def _remove_theme_font_attrs(rFonts) -> None:
    """移除 rFonts 中的主题字体属性，使显式字体（ascii/eastAsia 等）生效"""
    for attr in _THEME_FONT_ATTRS:
        if attr in rFonts.attrib:
            del rFonts.attrib[attr]


def _apply_font_to_style(style, font_config: FontConfig) -> None:
    """
    将 FontConfig 应用到 Word 样式定义（修改样式本身，而非直接格式）
    需同时设置 ascii 和 eastAsia 以支持中文字体。
    Heading 等样式使用主题字体，会覆盖显式字体，故需先移除主题属性。
    """
    font_ascii = _resolve_font(font_config.name_ascii or font_config.name)
    font_east_asia = _resolve_font(font_config.name_east_asia or font_config.name)
    rgb_color = _parse_color(font_config.color)

    font = style.font
    if font_ascii or font_east_asia:
        rPr = style._element.get_or_add_rPr()
        rFonts = rPr.get_or_add_rFonts()
        _remove_theme_font_attrs(rFonts)  # 移除主题字体，使显式字体生效
    if font_ascii:
        font.name = font_ascii
    if font_east_asia:
        try:
            rPr = style._element.get_or_add_rPr()
            rFonts = rPr.get_or_add_rFonts()
            rFonts.set(qn("w:eastAsia"), font_east_asia)
        except Exception:
            font.name = font_east_asia
    if font_config.size_pt is not None:
        font.size = Pt(font_config.size_pt)
    if rgb_color:
        font.color.rgb = rgb_color
    if font_config.bold is not None:
        font.bold = font_config.bold
    if font_config.italic is not None:
        font.italic = font_config.italic
    if font_config.underline is not None:
        font.underline = font_config.underline


def _apply_font_to_run(run, font_config: FontConfig) -> None:
    """
    将 FontConfig 应用到 run 的直接格式（批量设置已有内容的字体）
    """
    font_ascii = _resolve_font(font_config.name_ascii or font_config.name)
    font_east_asia = _resolve_font(font_config.name_east_asia or font_config.name)
    rgb_color = _parse_color(font_config.color)

    if font_ascii:
        run.font.name = font_ascii
    if font_east_asia:
        try:
            rPr = run._element.get_or_add_rPr()
            rFonts = rPr.get_or_add_rFonts()
            rFonts.set(qn("w:eastAsia"), font_east_asia)
        except Exception:
            run.font.name = font_east_asia
    if font_config.size_pt is not None:
        run.font.size = Pt(font_config.size_pt)
    if rgb_color:
        run.font.color.rgb = rgb_color
    if font_config.bold is not None:
        run.font.bold = font_config.bold
    if font_config.italic is not None:
        run.font.italic = font_config.italic
    if font_config.underline is not None:
        run.font.underline = font_config.underline


class WordProcessor:
    """
    Word 文档排版处理器

    封装 python-docx 的文档操作，提供批量样式设置、文本替换、段落整形等功能。
    """

    def __init__(self, document: Document):
        """
        初始化排版处理器

        Args:
            document: python-docx Document 对象
        """
        self._doc = document

    @classmethod
    def from_path(cls, path: str) -> "WordProcessor":
        """
        从文件路径加载文档并创建 WordProcessor

        Args:
            path: .docx 文件路径

        Returns:
            WordProcessor 实例
        """
        logger.debug("加载文档: %s", path)
        doc = Document(path)
        return cls(doc)

    def save(self, path: Optional[str] = None) -> str:
        """
        保存文档

        Args:
            path: 保存路径，若为 None 则覆盖原文件（需在 from_path 加载时有效）

        Returns:
            保存后的文件路径
        """
        if path:
            logger.debug("保存文档至: %s", path)
            self._doc.save(path)
            return path
        raise ValueError("保存路径不能为空")

    def set_global_styles(
        self,
        font_config: FontConfig,
        target_scope: str = "all",
    ) -> None:
        """
        同时执行：1) 修改 Word 文档的样式定义；2) 批量设置已有内容的字体直接格式。

        - 样式定义：Normal、Heading 1 等会更新，新输入内容将自动继承
        - 批量设置：遍历匹配范围的段落/runs，直接应用字体格式，确保已有内容立即生效

        Args:
            font_config: 字体配置
            target_scope: 作用范围。支持：body、heading、heading_1~heading_9、caption、all
        """
        # 1. 修改样式定义
        style_names = _get_style_names_for_scope(target_scope)
        logger.info(
            "set_global_styles: 修改样式定义，范围=%s, 样式=%s",
            target_scope,
            style_names,
        )
        for style_name in style_names:
            try:
                style = self._doc.styles[style_name]
                _apply_font_to_style(style, font_config)
            except KeyError:
                logger.debug("样式 %s 不存在于文档中，跳过", style_name)

        # 2. 批量设置已有内容的字体直接格式
        run_count = 0
        for para in self._doc.paragraphs:
            if not self._match_scope(para, target_scope):
                continue
            for run in para.runs:
                _apply_font_to_run(run, font_config)
                run_count += 1
        logger.info("set_global_styles 完成：样式定义已更新，共处理 %d 个 run", run_count)

    def semantic_replace(
        self,
        search: str,
        replace: str,
        use_regex: bool = False,
    ) -> int:
        """
        智能文本替换（支持正则表达式）

        Args:
            search: 查找内容
            replace: 替换内容
            use_regex: 是否将 search 视为正则表达式

        Returns:
            替换次数
        """
        count = 0
        for para in self._doc.paragraphs:
            for run in para.runs:
                text = run.text
                if not text:
                    continue
                if use_regex:
                    new_text, n = re.subn(search, replace, text)
                else:
                    new_text = text.replace(search, replace)
                    n = text.count(search)
                if n > 0:
                    run.text = new_text
                    count += n

        # 同时处理表格中的文本
        for table in self._doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            text = run.text
                            if not text:
                                continue
                            if use_regex:
                                new_text, n = re.subn(search, replace, text)
                            else:
                                new_text = text.replace(search, replace)
                                n = text.count(search)
                            if n > 0:
                                run.text = new_text
                                count += n

        logger.info("semantic_replace 完成: %r -> %r, 共替换 %d 处", search, replace, count)
        return count

    def paragraph_shaping(
        self,
        paragraph_config: ParagraphConfig,
        target_scope: str = "all",
    ) -> None:
        """
        同时执行：1) 修改 Word 文档样式的段落格式定义；2) 批量设置已有段落的直接格式。

        - 样式定义：新输入的内容将自动继承行距、首行缩进等
        - 批量设置：遍历匹配范围的段落，直接应用段落格式，确保已有内容立即生效

        Args:
            paragraph_config: 段落配置
            target_scope: 作用范围
        """
        # 1. 修改样式定义
        style_names = _get_style_names_for_scope(target_scope)
        logger.info(
            "paragraph_shaping: 修改样式定义，范围=%s, 样式=%s, 首行缩进=%s, 行距=%s",
            target_scope,
            style_names,
            paragraph_config.first_line_indent,
            paragraph_config.line_spacing,
        )
        for style_name in style_names:
            try:
                style = self._doc.styles[style_name]
                pf = style.paragraph_format
                if paragraph_config.first_line_indent is not None:
                    pf.first_line_indent = Pt(paragraph_config.first_line_indent)
                if paragraph_config.line_spacing is not None:
                    pf.line_spacing = paragraph_config.line_spacing
                if paragraph_config.space_before is not None:
                    pf.space_before = Pt(paragraph_config.space_before)
                if paragraph_config.space_after is not None:
                    pf.space_after = Pt(paragraph_config.space_after)
                align = _alignment_from_str(paragraph_config.alignment)
                if align is not None:
                    pf.alignment = align
            except KeyError:
                logger.debug("样式 %s 不存在于文档中，跳过", style_name)

        # 2. 批量设置已有段落的直接格式
        para_count = 0
        for para in self._doc.paragraphs:
            if not self._match_scope(para, target_scope):
                continue
            pf = para.paragraph_format
            if paragraph_config.first_line_indent is not None:
                pf.first_line_indent = Pt(paragraph_config.first_line_indent)
            if paragraph_config.line_spacing is not None:
                pf.line_spacing = paragraph_config.line_spacing
            if paragraph_config.space_before is not None:
                pf.space_before = Pt(paragraph_config.space_before)
            if paragraph_config.space_after is not None:
                pf.space_after = Pt(paragraph_config.space_after)
            align = _alignment_from_str(paragraph_config.alignment)
            if align is not None:
                pf.alignment = align
            para_count += 1
        logger.info("paragraph_shaping 完成：样式定义已更新，共处理 %d 个段落", para_count)

    def apply_first_line_indent(
        self,
        indent_pt: float,
        target_scope: str = "body",
    ) -> None:
        """
        修改样式的首行缩进定义

        Args:
            indent_pt: 缩进量（磅），2 字符 ≈ 24pt
            target_scope: 作用范围（body、heading、heading_1~9、caption、all）
        """
        self.paragraph_shaping(
            ParagraphConfig(first_line_indent=indent_pt),
            target_scope=target_scope,
        )

    def _is_heading(self, paragraph) -> bool:
        """判断段落是否为标题样式"""
        return self._get_heading_level(paragraph) is not None

    def _get_heading_level(self, paragraph) -> Optional[int]:
        """
        获取段落的标题级别（1-9），非标题返回 None
        """
        if paragraph.style is None:
            return None
        style_name = getattr(paragraph.style, "name", "") or ""
        if "heading" not in style_name.lower() and "标题" not in style_name:
            return None
        # 匹配 Heading 1, 标题 1, Heading1 等
        m = re.search(r"(?:heading|标题)\s*(\d)", style_name, re.IGNORECASE)
        if m:
            return int(m.group(1))
        return 1  # 未明确级别时默认 1

    def _is_caption(self, paragraph) -> bool:
        """判断段落是否为题注样式"""
        if paragraph.style is None:
            return False
        style_name = getattr(paragraph.style, "name", "") or ""
        name_lower = style_name.lower()
        return "caption" in name_lower or "题注" in style_name

    def _match_scope(self, paragraph, target_scope: str) -> bool:
        """
        判断段落是否匹配目标作用范围

        支持：body、heading、heading_1~heading_9、caption、all
        """
        if target_scope == "all":
            return True
        if target_scope == "body":
            return not self._is_heading(paragraph) and not self._is_caption(paragraph)
        if target_scope == "heading":
            return self._is_heading(paragraph)
        if target_scope.startswith("heading_"):
            try:
                level = int(target_scope.split("_")[1])
                return self._get_heading_level(paragraph) == level
            except (ValueError, IndexError):
                return False
        if target_scope == "caption":
            return self._is_caption(paragraph)
        if target_scope == "emphasis":
            # emphasis 为 run 级，段落级视为不匹配（由后续扩展处理）
            return False
        return True

    @property
    def document(self) -> Document:
        """获取底层 Document 对象"""
        return self._doc
