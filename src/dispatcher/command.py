"""
命令调度器

将 TypesettingIntent 映射到 WordProcessor 执行。
统一使用 intent.to_scope_rules() 获取待应用的规则列表。
"""

import logging

from ..schemas.typesetting import TypesettingIntent
from ..formatter.engine import WordProcessor

logger = logging.getLogger(__name__)


class CommandDispatcher:
    """
    命令调度器

    根据 TypesettingIntent 调用 WordProcessor 完成排版。
    通过 intent.to_scope_rules() 统一为 scope_rules 流程。
    """

    @staticmethod
    def dispatch(intent: TypesettingIntent, word_processor: WordProcessor) -> None:
        """
        将排版意图调度到 WordProcessor 执行

        使用 intent.to_scope_rules() 统一为 scope_rules 格式后逐条应用。

        Args:
            intent: 解析后的排版意图（Schema 类型）
            word_processor: Word 排版处理器实例
        """
        scope_rules = intent.to_scope_rules()
        if not scope_rules:
            logger.info("无待应用的 scope 规则，跳过样式设置")
            return

        for sr in scope_rules:
            scope = sr.target_scope
            logger.info("调度 scope_rules，作用范围: %s", scope)
            if sr.font_config is not None:
                gs = sr.font_config
                font_desc = (
                    f"西文={gs.name_ascii or gs.name}, 中文={gs.name_east_asia or gs.name}"
                    if (gs.name_ascii or gs.name_east_asia)
                    else f"字体={gs.name}"
                )
                logger.info("执行 set_global_styles: %s, 字号=%s", font_desc, gs.size_pt)
                word_processor.set_global_styles(
                    font_config=sr.font_config,
                    target_scope=scope,
                )
            if sr.paragraph_config is not None:
                logger.info("执行 paragraph_shaping: 首行缩进=%s, 行距=%s", sr.paragraph_config.first_line_indent, sr.paragraph_config.line_spacing)
                word_processor.paragraph_shaping(
                    paragraph_config=sr.paragraph_config,
                    target_scope=scope,
                )

        # 文本替换
        for rule in intent.replacements:
            logger.info("执行 semantic_replace: %r -> %r (regex=%s)", rule.search, rule.replace, rule.use_regex)
            word_processor.semantic_replace(
                search=rule.search,
                replace=rule.replace,
                use_regex=rule.use_regex,
            )
