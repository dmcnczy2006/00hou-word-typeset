"""
智能 Word 排版 Agent 主入口

提供 process_document 作为主入口函数，完成：解析用户意图 → 调度排版引擎 → 保存文档。
"""

import logging
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv

load_dotenv()  # 加载 .env 中的 OPENAI_API_KEY 等配置

logger = logging.getLogger(__name__)

from .config.presets import Config
from .dispatcher.command import CommandDispatcher
from .formatter.engine import WordProcessor
from .intent.parser import IntentParser


def process_document(
    file_path: str,
    user_prompt: str,
    preset: str = "official",
    output_path: Optional[str] = None,
    llm_connector=None,
) -> str:
    """
    主入口：解析用户意图 → 调度排版引擎 → 保存文档

    Args:
        file_path: 待处理的 .docx 文件路径
        user_prompt: 用户自然语言排版指令，如「把正文设为小四号仿宋，首行缩进2字符」
        preset: 预设规则名称，如 "official"、"thesis"、"default"
        output_path: 输出文件路径，若为 None 则覆盖原文件
        llm_connector: LLM 连接器实例，若为 None 则使用 OpenAIConnector（从 .env 读取 API Key）

    Returns:
        处理后的文件路径

    Example:
        >>> process_document("report.docx", "正文小四仿宋，首行缩进2字符", preset="official")
        'report.docx'
    """
    # 1. 加载文档
    path = Path(file_path)
    logger.info("开始处理文档: %s", file_path)
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")
    if path.suffix.lower() != ".docx":
        raise ValueError("仅支持 .docx 格式")

    word_processor = WordProcessor.from_path(str(path))
    logger.info("文档加载成功，共 %d 个段落", len(word_processor.document.paragraphs))

    # 2. 解析用户意图（与预设合并）
    logger.info("解析用户意图，预设=%s，指令=%s", preset, user_prompt[:50] + "..." if len(user_prompt) > 50 else user_prompt)
    parser = IntentParser(llm_connector=llm_connector)
    intent = parser.parse(
        user_prompt=user_prompt,
        preset=preset,
        merge_with_preset=True,
    )
    logger.info("意图解析完成")

    # 3. 调度排版引擎执行
    logger.info("调度排版引擎执行")
    CommandDispatcher.dispatch(intent, word_processor)

    # 4. 保存
    save_path = output_path or str(path)
    word_processor.save(save_path)
    logger.info("文档已保存至: %s", save_path)

    return save_path


def create_sample_document(path: str = "sample.docx") -> str:
    """
    创建示例文档（用于测试）

    Args:
        path: 保存路径

    Returns:
        保存后的文件路径
    """
    from docx import Document

    doc = Document()
    doc.add_heading("示例文档", 0)
    doc.add_paragraph(
        "这是一段正文。智能 Word 排版 Agent 可以根据自然语言指令自动调整字体、字号、段落格式等。"
    )
    doc.add_paragraph(
        "例如：「把正文设为小四号仿宋，首行缩进2字符」即可一键应用公文标准格式。"
    )
    doc.save(path)
    return path


if __name__ == "__main__":
    # 命令行入口示例
    import sys

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )

    if len(sys.argv) < 2:
        print("用法: python -m src.main <docx路径> [用户指令]")
        print("示例: python -m src.main sample.docx \"正文小四仿宋，首行缩进2字符\"")
        sys.exit(1)

    doc_path = sys.argv[1]
    prompt = sys.argv[2] if len(sys.argv) > 2 else ""

    # 若文件不存在则创建示例
    if not Path(doc_path).exists():
        create_sample_document(doc_path)
        print(f"已创建示例文档: {doc_path}")

    result = process_document(doc_path, prompt, preset="official")
    print(f"处理完成，已保存至: {result}")
