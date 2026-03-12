"""
Z Mode 工作流解析提示词

将用户自然语言描述的多步骤工作流解析为结构化 Workflow JSON。

调用关系：
    intent.workflow_parser.WorkflowParser.parse()
        → build_workflow_prompt(user_prompt, schema_json)
            → 返回完整 prompt
"""

from typing import Optional

WORKFLOW_PROMPT_TEMPLATE = """你是一个 Word 文档排版工作流解析助手。用户会描述一个多步骤的排版工作流，请将其解析为结构化的 JSON。

## 用户描述的工作流
{user_prompt}

## 输出要求
请严格按照以下 JSON Schema 输出，只返回 JSON 对象，不要包含 markdown 代码块或其它说明文字。

### 步骤类型说明

1. **heading_change**：标题升降级
   - 将某级标题改为另一级，如「标题2变为标题1」→ {{"type": "heading_change", "from_scope": "heading_2", "to_scope": "heading_1"}}
   - from_scope/to_scope 取值：heading_1 ~ heading_9

2. **apply_list**：应用多级列表
   - 为文档的标题或正文设置多级列表
   - target_scope：heading（仅标题）、body（仅正文）、all（全文）
     * 【重要】多级列表、1.1.1、一（一）1（1）等章节编号格式，默认用 heading（仅标题），正文不应被编号
     * 只有用户明确说「正文」「全文」「所有段落」「每段都编号」时才用 body 或 all
   - list_style：根据用户自然语言映射
     * decimal：1, 2, 3（各级独立）
     * decimal_hierarchical：1, 1.1, 1.1.1（层级编号）
     * chinese_mixed：一、（一）、1、（1）（公文/法律常用）
     * bullet：项目符号
     * multilevel：默认多级
   - 如「1.1.1 格式」「一级用一二级用（一）」「一（一）1（1）」→ 对应上述值，且 target_scope 用 heading

3. **re_typeset**：再排版
   - 完成一次字体、段落格式的再排版
   - preset：预设名称，如 official、thesis、default
   - user_override：用户对字体/字号的补充描述，如「正文小四号，标题一号黑体」，若无则填空字符串

4. **page_margins**：页面边距设置
   - top_pt、bottom_pt、left_pt、right_pt：上下左右边距（磅），1英寸≈72pt，仅指定用户提到的
   - 如「上下左右各2.5cm」→ 2.5cm≈72pt，四个都设；「左边距3cm」→ 仅 left_pt: 85

5. **header_footer**：页眉页脚设置。用户不说则无，说得模糊则用默认。
   - header_text、footer_text：页眉/页脚文字（odd_even 时作为奇数页默认）
   - odd_even：奇偶页不同（默认 false）
   - odd_header_text、even_header_text：奇偶页分别的页眉文字
   - odd_footer_text、even_footer_text：奇偶页分别的页脚文字
   - show_chapter：显示当前章节（默认 false）；odd_show_chapter、even_show_chapter 可分别指定
   - underline：页眉/页脚下划线；odd_underline、even_underline 可分别指定
   - page_number：添加页码；odd_page_number、even_page_number 可分别指定
   - first_page_different：首页不同（默认 false）
   - 如「奇偶页不同，奇数页显示文档名、偶数页显示章节」→ odd_even: true, odd_header_text: "文档名", even_show_chapter: true

6. **update_fields**：更新域
   - 无参数，设置文档打开时自动更新域（如页码、目录等）

7. **collapse_empty_lines**：连续空行合并
   - from_count：连续空行数量 x（>=2）
   - to_count：目标空行数量 y（>=0）
   - 如「3个连续空行改成1个」→ {{"type": "collapse_empty_lines", "from_count": 3, "to_count": 1}}

### 解析规则
- 按用户描述的步骤顺序输出 steps 数组
- 若用户只说单一步骤（如仅再排版），也输出单元素数组
- 标题级别：heading_1 对应一级标题，heading_2 对应二级标题，以此类推
- apply_list 的 target_scope：用户说「多级列表」「1.1.1」「一（一）1（1）」等且未明确说正文/全文时，必须用 heading
- 若描述中包含「用户补充」，请在原工作流基础上按补充内容调整（修改、增删步骤）

## JSON Schema
{schema_json}

请直接输出 JSON："""


MODE_DETECTION_PROMPT = """你是一个 Word 文档排版模式判断助手。根据用户的排版指令，判断应使用哪种模式。

## 模式说明
- **Y 模式**：单次再排版。用户描述的是字体/段落/格式调整，如「正文小四仿宋」「首行缩进2字符」「标题一号黑体」中的一个或多个。一次意图即可完成。
- **Z 模式**：多步骤工作流。用户明确描述了多个独立步骤，如「1. 边距 2. 页码 3. 再排版」「先改标题级别，再设置列表，最后再排版」「设置边距；添加页码；更新域」等。需要按顺序执行多个不同性质的操作。

## 用户指令
{user_prompt}

## 输出要求
只输出一个字母：y 或 z。不要输出任何其它文字、标点或解释。
- 输出 y：单次再排版（Y 模式）
- 输出 z：多步骤工作流（Z 模式）

输出："""


def build_mode_detection_prompt(user_prompt: str) -> str:
    """构建模式判断的 LLM prompt"""
    return MODE_DETECTION_PROMPT.format(user_prompt=user_prompt.strip())


def build_workflow_prompt(
    user_prompt: str,
    schema_json: str,
) -> str:
    """
    构建工作流解析的完整 LLM prompt

    Args:
        user_prompt: 用户自然语言描述的工作流
        schema_json: Workflow 的 JSON Schema 字符串

    Returns:
        完整的 prompt 字符串
    """
    return WORKFLOW_PROMPT_TEMPLATE.format(
        user_prompt=user_prompt.strip(),
        schema_json=schema_json,
    )
