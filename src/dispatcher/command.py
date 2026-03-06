"""
命令调度器

将 IntentParser 解析出的 TypesettingIntent 映射到 WordProcessor 的具体方法执行。
"""

import logging

from ..schemas.typesetting import TypesettingIntent
from ..formatter.engine import WordProcessor

logger = logging.getLogger(__name__)


class CommandDispatcher:
    """
    命令调度器

    根据 TypesettingIntent 调用 WordProcessor 的相应方法完成排版。
    """

    @staticmethod
    def dispatch(intent: TypesettingIntent, word_processor: WordProcessor) -> None:
        """
        将排版意图调度到 WordProcessor 执行

        Args:
            intent: 解析后的排版意图
            word_processor: Word 排版处理器实例
        """
        target_scope = intent.target_scope or "all"
        logger.info("调度排版指令，作用范围: %s", target_scope)

        # 1. 全局样式
        if intent.global_styles is not None:
            logger.info("执行 set_global_styles: 字体=%s, 字号=%s", intent.global_styles.name, intent.global_styles.size_pt)
            word_processor.set_global_styles(
                font_config=intent.global_styles,
                target_scope=target_scope,
            )

        # 2. 段落整形
        if intent.paragraph_config is not None:
            logger.info("执行 paragraph_shaping: 首行缩进=%s, 行距=%s", intent.paragraph_config.first_line_indent, intent.paragraph_config.line_spacing)
            word_processor.paragraph_shaping(
                paragraph_config=intent.paragraph_config,
                target_scope=target_scope,
            )

        # 3. 文本替换
        for rule in intent.replacements:
            logger.info("执行 semantic_replace: %r -> %r (regex=%s)", rule.search, rule.replace, rule.use_regex)
            word_processor.semantic_replace(
                search=rule.search,
                replace=rule.replace,
                use_regex=rule.use_regex,
            )
