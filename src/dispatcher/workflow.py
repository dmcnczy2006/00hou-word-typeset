"""
工作流调度器（Z Mode）

按序执行 Workflow 中的各步骤，提供工作流预览格式化。
"""

import logging
from typing import Optional

from ..schemas.typesetting import (
    ApplyListStep,
    CollapseEmptyLinesStep,
    HeaderFooterStep,
    HeadingChangeStep,
    PageMarginsStep,
    ReTypesetStep,
    UpdateFieldsStep,
    Workflow,
)
from ..formatter.engine import WordProcessor
from ..intent.parser import IntentParser
from .command import CommandDispatcher

logger = logging.getLogger(__name__)


class WorkflowDispatcher:
    """
    工作流调度器

    根据 Workflow 按序执行各步骤，调用 WordProcessor 和 CommandDispatcher。
    """

    @staticmethod
    def execute(
        workflow: Workflow,
        word_processor: WordProcessor,
        llm_connector=None,
    ) -> None:
        """
        执行工作流中的各步骤

        Args:
            workflow: 解析后的工作流
            word_processor: Word 排版处理器实例
            llm_connector: LLM 连接器，用于 re_typeset 步骤的意图解析
        """
        for i, step in enumerate(workflow.steps):
            logger.info("执行工作流步骤 %d/%d: type=%s", i + 1, len(workflow.steps), step.type)
            if isinstance(step, HeadingChangeStep):
                word_processor.change_heading_level(
                    from_scope=step.from_scope,
                    to_scope=step.to_scope,
                )
            elif isinstance(step, ApplyListStep):
                word_processor.apply_multilevel_list(
                    target_scope=step.target_scope,
                    list_style=step.list_style,
                )
            elif isinstance(step, ReTypesetStep):
                intent_parser = IntentParser(llm_connector=llm_connector)
                intent = intent_parser.parse(
                    user_prompt=step.user_override or "",
                    preset=step.preset or "official",
                    merge_with_preset=True,
                )
                CommandDispatcher.dispatch(intent, word_processor)
            elif isinstance(step, PageMarginsStep):
                word_processor.set_page_margins(
                    top_pt=step.top_pt,
                    bottom_pt=step.bottom_pt,
                    left_pt=step.left_pt,
                    right_pt=step.right_pt,
                )
            elif isinstance(step, HeaderFooterStep):
                word_processor.set_header_footer(
                    header_text=step.header_text,
                    footer_text=step.footer_text,
                    odd_even=step.odd_even,
                    odd_header_text=step.odd_header_text,
                    even_header_text=step.even_header_text,
                    odd_footer_text=step.odd_footer_text,
                    even_footer_text=step.even_footer_text,
                    show_chapter=step.show_chapter,
                    odd_show_chapter=step.odd_show_chapter,
                    even_show_chapter=step.even_show_chapter,
                    underline=step.underline,
                    odd_underline=step.odd_underline,
                    even_underline=step.even_underline,
                    page_number=step.page_number,
                    odd_page_number=step.odd_page_number,
                    even_page_number=step.even_page_number,
                    first_page_different=step.first_page_different,
                )
            elif isinstance(step, UpdateFieldsStep):
                word_processor.update_fields()
            elif isinstance(step, CollapseEmptyLinesStep):
                word_processor.collapse_empty_lines(
                    from_count=step.from_count,
                    to_count=step.to_count,
                )
            else:
                logger.warning("未知步骤类型: %s", getattr(step, "type", step))


def format_workflow_for_display(workflow: Workflow) -> str:
    """
    将工作流格式化为可读文本，供用户确认前展示。

    Args:
        workflow: 解析后的工作流

    Returns:
        格式化后的字符串
    """
    lines = [f"工作流预览（共 {len(workflow.steps)} 步）：", ""]
    for i, step in enumerate(workflow.steps, 1):
        desc = _format_step(step)
        lines.append(f"  {i}. {desc}")
    return "\n".join(lines)


def _format_step(step) -> str:
    """格式化单步为可读描述"""
    t = step.type
    if t == "heading_change":
        return f"标题升降级：{step.from_scope} → {step.to_scope}"
    if t == "apply_list":
        style_desc = {
            "decimal": "1, 2, 3",
            "decimal_hierarchical": "1, 1.1, 1.1.1",
            "chinese_mixed": "一、（一）、1、（1）",
            "bullet": "项目符号",
            "multilevel": "默认多级",
        }.get(step.list_style, step.list_style)
        return f"应用列表：范围={step.target_scope}，样式={style_desc}"
    if t == "re_typeset":
        ov = step.user_override or "（无）"
        return f"再排版：预设={step.preset or 'official'}，补充={ov}"
    if t == "page_margins":
        parts = []
        if step.top_pt is not None:
            parts.append(f"上{step.top_pt}pt")
        if step.bottom_pt is not None:
            parts.append(f"下{step.bottom_pt}pt")
        if step.left_pt is not None:
            parts.append(f"左{step.left_pt}pt")
        if step.right_pt is not None:
            parts.append(f"右{step.right_pt}pt")
        return f"页面边距：{', '.join(parts) or '无'}"
    if t == "header_footer":
        parts = []
        if step.header_text or step.odd_header_text:
            parts.append(f"页眉={step.odd_header_text or step.header_text or ''}")
        if step.even_header_text:
            parts.append(f"偶数页眉={step.even_header_text}")
        if step.footer_text or step.odd_footer_text:
            parts.append(f"页脚={step.odd_footer_text or step.footer_text or ''}")
        if step.even_footer_text:
            parts.append(f"偶数页脚={step.even_footer_text}")
        if step.page_number or step.odd_page_number or step.even_page_number:
            parts.append("页码")
        if step.show_chapter or step.odd_show_chapter or step.even_show_chapter:
            parts.append("显示章节")
        if step.underline or step.odd_underline or step.even_underline:
            parts.append("下划线")
        if step.odd_even:
            parts.append("奇偶页不同")
        return f"页眉页脚：{', '.join(parts) or '无'}"
    if t == "update_fields":
        return "更新域"
    if t == "collapse_empty_lines":
        return f"连续空行：{step.from_count}个→{step.to_count}个"
    return str(step)
