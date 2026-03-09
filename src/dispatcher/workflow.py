"""
工作流调度器（Z Mode）

按序执行 Workflow 中的各步骤。
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
