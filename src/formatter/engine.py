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
        批量设置全局字体样式（字体、字号、颜色、对齐）

        Args:
            font_config: 字体配置
            target_scope: 作用范围，"body"=正文段落，"heading"=标题，"all"=全文
        """
        # 解析字体名（智能映射）：中西文可分别指定
        def _resolve_font(raw: Optional[str]) -> Optional[str]:
            if not raw:
                return None
            return FONT_NAME_MAP.get(raw, raw)

        font_ascii = _resolve_font(font_config.name_ascii or font_config.name)
        font_east_asia = _resolve_font(font_config.name_east_asia or font_config.name)

        # 解析颜色
        rgb_color = _parse_color(font_config.color)

        logger.info(
            "set_global_styles: 西文=%s, 中文=%s, 字号=%spt, 范围=%s",
            font_ascii,
            font_east_asia,
            font_config.size_pt,
            target_scope,
        )
        affected = 0
        for para in self._doc.paragraphs:
            # 根据 target_scope 过滤
            if target_scope == "heading":
                if not self._is_heading(para):
                    continue
            elif target_scope == "body":
                if self._is_heading(para):
                    continue

            for run in para.runs:
                # 西文：w:ascii / w:hAnsi（run.font.name 会创建 rPr/rFonts）
                if font_ascii:
                    run.font.name = font_ascii
                # 中文：w:eastAsia（需单独设置，否则 CJK 不生效）
                if font_east_asia:
                    try:
                        rPr = run._element.get_or_add_rPr()
                        if rPr.rFonts is not None:
                            rPr.rFonts.set(qn("w:eastAsia"), font_east_asia)
                        else:
                            run.font.name = font_east_asia
                            rPr = run._element.get_or_add_rPr()
                            if rPr.rFonts is not None:
                                rPr.rFonts.set(qn("w:eastAsia"), font_east_asia)
                    except Exception:
                        run.font.name = font_east_asia
                if font_config.size_pt is not None:
                    run.font.size = Pt(font_config.size_pt)
                if rgb_color:
                    run.font.color.rgb = rgb_color
                affected += 1
        logger.info("set_global_styles 完成，共处理 %d 个 run", affected)

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
        段落整形：设置行间距、段前段后距离、对齐方式

        Args:
            paragraph_config: 段落配置
            target_scope: 作用范围
        """
        logger.info("paragraph_shaping: 范围=%s, 首行缩进=%s, 行距=%s", target_scope, paragraph_config.first_line_indent, paragraph_config.line_spacing)
        affected = 0
        for para in self._doc.paragraphs:
            if target_scope == "heading" and not self._is_heading(para):
                continue
            if target_scope == "body" and self._is_heading(para):
                continue

            pf = para.paragraph_format

            if paragraph_config.first_line_indent is not None:
                pf.first_line_indent = Pt(paragraph_config.first_line_indent)
            if paragraph_config.line_spacing is not None:
                # 倍行距：1.5 表示 1.5 倍
                pf.line_spacing = paragraph_config.line_spacing
            if paragraph_config.space_before is not None:
                pf.space_before = Pt(paragraph_config.space_before)
            if paragraph_config.space_after is not None:
                pf.space_after = Pt(paragraph_config.space_after)
            align = _alignment_from_str(paragraph_config.alignment)
            if align is not None:
                pf.alignment = align
            affected += 1
        logger.info("paragraph_shaping 完成，共处理 %d 个段落", affected)

    def apply_first_line_indent(
        self,
        indent_pt: float,
        target_scope: str = "body",
    ) -> None:
        """
        应用首行缩进

        Args:
            indent_pt: 缩进量（磅），2 字符 ≈ 24pt
            target_scope: 作用范围
        """
        for para in self._doc.paragraphs:
            if target_scope == "heading" and not self._is_heading(para):
                continue
            if target_scope == "body" and self._is_heading(para):
                continue
            para.paragraph_format.first_line_indent = Pt(indent_pt)

    def _is_heading(self, paragraph) -> bool:
        """判断段落是否为标题样式"""
        if paragraph.style is None:
            return False
        style_name = getattr(paragraph.style, "name", "") or ""
        return "heading" in style_name.lower() or "标题" in style_name

    @property
    def document(self) -> Document:
        """获取底层 Document 对象"""
        return self._doc
