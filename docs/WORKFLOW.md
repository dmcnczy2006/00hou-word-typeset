# Z Mode 工作流文档

Z Mode 支持用户通过自然语言描述多步骤排版工作流，由 LLM 解析为结构化 JSON，按序执行各步骤。

## 概述

| 模式 | 说明 |
|------|------|
| **Y Mode** | 单次再排版，用户指令 → TypesettingIntent → 一次执行 |
| **Z Mode** | 多步骤工作流，用户描述 → Workflow JSON → 按序执行各步骤 |

用户可通过 `mode` 参数选择：`y`（强制 Y）、`z`（强制 Z）、`auto`（根据指令自动判断）。

## 工作流结构

```json
{
  "mode": "z",
  "steps": [
    { "type": "heading_change", "from_scope": "heading_2", "to_scope": "heading_1" },
    { "type": "apply_list", "target_scope": "heading", "list_style": "multilevel" },
    { "type": "re_typeset", "preset": "official", "user_override": "正文小四、标题一号" }
  ]
}
```

- `mode`：固定为 `"z"`
- `steps`：步骤数组，按顺序执行
- 步骤采用 **discriminated union**，通过 `type` 字段区分

## 步骤类型

### 1. heading_change（标题升降级）

将指定级别的标题改为另一级别。

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| from_scope | string | 是 | 源 scope，如 `heading_2` |
| to_scope | string | 是 | 目标 scope，如 `heading_1` |

**scope 取值**：`heading_1` ~ `heading_9`

**示例**：
```json
{ "type": "heading_change", "from_scope": "heading_2", "to_scope": "heading_1" }
```

---

### 2. apply_list（应用列表）

为文档的标题或正文设置多级列表。

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| target_scope | string | 是 | `heading`（仅标题）、`body`（仅正文）、`all`（全文） |
| list_style | string | 是 | `multilevel`、`decimal`、`bullet` |

**示例**：
```json
{ "type": "apply_list", "target_scope": "heading", "list_style": "multilevel" }
```

---

### 3. re_typeset（再排版）

完成一次字体、段落格式的再排版，复用 Y Mode 的 IntentParser。

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| preset | string | 否 | 预设名称，默认 `official` |
| user_override | string | 否 | 自然语言补充，如「正文小四、标题一号」 |

**示例**：
```json
{ "type": "re_typeset", "preset": "official", "user_override": "正文小四号，标题一号黑体" }
```

---

### 4. page_margins（页面边距）

设置页面边距。仅指定用户提到的边距，未指定的保持不变。

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| top_pt | float | 否 | 上边距（磅），1 英寸 ≈ 72pt |
| bottom_pt | float | 否 | 下边距 |
| left_pt | float | 否 | 左边距 |
| right_pt | float | 否 | 右边距 |

**单位换算**：2.5cm ≈ 72pt，3cm ≈ 85pt

**示例**：
```json
{ "type": "page_margins", "top_pt": 72, "bottom_pt": 72, "left_pt": 72, "right_pt": 72 }
```

---

### 5. header_footer（页眉页脚）

设置页眉页脚。**用户不说则无，说得模糊则用默认**。

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| header_text | string | 否 | 页眉文字（odd_even 时作为奇数页默认） |
| footer_text | string | 否 | 页脚文字（odd_even 时作为奇数页默认） |
| odd_even | bool | 否 | 奇偶页不同，默认 `false` |
| odd_header_text | string | 否 | 奇数页页眉，odd_even 时优先于 header_text |
| even_header_text | string | 否 | 偶数页页眉 |
| odd_footer_text | string | 否 | 奇数页页脚 |
| even_footer_text | string | 否 | 偶数页页脚 |
| show_chapter | bool | 否 | 显示当前章节，默认 `false` |
| odd_show_chapter | bool | 否 | 奇数页显示章节 |
| even_show_chapter | bool | 否 | 偶数页显示章节 |
| underline | bool | 否 | 页眉/页脚下划线 |
| odd_underline | bool | 否 | 奇数页下划线 |
| even_underline | bool | 否 | 偶数页下划线 |
| page_number | bool | 否 | 添加页码 |
| odd_page_number | bool | 否 | 奇数页页码 |
| even_page_number | bool | 否 | 偶数页页码 |
| first_page_different | bool | 否 | 首页不同，默认 `false` |

**示例**：
```json
{ "type": "header_footer", "page_number": true }
{ "type": "header_footer", "odd_even": true, "odd_header_text": "文档标题", "even_show_chapter": true }
```

---

### 6. update_fields（更新域）

设置文档在打开时自动更新域（如页码、目录）。无参数。

**示例**：
```json
{ "type": "update_fields" }
```

---

### 7. collapse_empty_lines（连续空行合并）

将 x 个连续空行改成 y 个空行。

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| from_count | int | 是 | 连续空行数量 x（≥2） |
| to_count | int | 是 | 目标空行数量 y（≥0） |

**示例**：
```json
{ "type": "collapse_empty_lines", "from_count": 3, "to_count": 1 }
```

---

## 数据流

```
用户自然语言
    ↓
WorkflowParser.parse()  [LLM 解析 → Workflow JSON]
    ↓
WorkflowDispatcher.execute()
    ↓
逐步骤执行
    ├─ heading_change    → WordProcessor.change_heading_level()
    ├─ apply_list        → WordProcessor.apply_multilevel_list()
    ├─ re_typeset        → IntentParser + CommandDispatcher
    ├─ page_margins      → WordProcessor.set_page_margins()
    ├─ header_footer     → WordProcessor.set_header_footer()
    ├─ update_fields     → WordProcessor.update_fields()
    └─ collapse_empty_lines → WordProcessor.collapse_empty_lines()
```

## 使用方式

### Python API

```python
from src.main import process_document

# Z Mode：多步骤工作流
process_document(
    file_path="report.docx",
    user_prompt="1. 将标题2变为标题1；2. 设置多级列表；3. 再排版，正文小四、标题一号",
    mode="z",
)

# 强制 Z Mode
process_document("report.docx", "添加页码；3个空行改1个", mode="z")
```

### 命令行

```bash
# Z Mode（-m z 强制指定）
python -m src.main sample.docx "1. 标题2变标题1 2. 多级列表 3. 再排版" -m z

# 指定输出路径
python -m src.main sample.docx "设置边距2.5cm；添加页码" -m z -o output.docx
```

### 自动模式判断

`mode="auto"` 时，以下情况会走 Z Mode：

- 包含编号（如 `1. xxx 2. xxx`）
- 分号分隔多句
- 含「步骤」「工作流」等词

## 自然语言示例

| 用户描述 | 解析结果（示意） |
|----------|------------------|
| 将文档里的所有标题2变为标题1 | `heading_change` from heading_2 to heading_1 |
| 为文档设置多级列表 | `apply_list` target_scope: heading |
| 完成一次再排版，正文小四，标题一号 | `re_typeset` user_override: "正文小四，标题一号" |
| 上下左右边距各2.5cm | `page_margins` 四边 72pt |
| 添加页码 | `header_footer` page_number: true |
| 页眉显示当前章节、下划线 | `header_footer` show_chapter: true, underline: true |
| 更新域 | `update_fields` |
| 3个连续空行改成1个 | `collapse_empty_lines` from_count: 3, to_count: 1 |

## 相关文件

| 文件 | 说明 |
|------|------|
| [src/schemas/typesetting.py](../src/schemas/typesetting.py) | Workflow、各 Step 的 Schema 定义 |
| [src/intent/workflow_parser.py](../src/intent/workflow_parser.py) | WorkflowParser 解析自然语言 |
| [src/prompts/workflow_prompt.py](../src/prompts/workflow_prompt.py) | LLM 提示词模板 |
| [src/dispatcher/workflow.py](../src/dispatcher/workflow.py) | WorkflowDispatcher 调度执行 |
| [src/formatter/engine.py](../src/formatter/engine.py) | WordProcessor 各步骤实现 |

## 扩展指南

新增步骤类型时：

1. 在 `typesetting.py` 中定义新的 Step 类（含 `type: Literal["xxx"]`）
2. 将新 Step 加入 `WorkflowStep` 的 Union
3. 在 `engine.py` 中实现对应方法
4. 在 `workflow.py` 的 `WorkflowDispatcher.execute()` 中增加分支
5. 在 `workflow_prompt.py` 中补充步骤说明
