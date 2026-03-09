# 智能 Word 排版 Agent

基于 python-docx、LLM 和 Pydantic 的智能 Word 文档排版工具，支持自然语言指令驱动排版。

## 功能特性

- **Y Mode（单次再排版）**：将自然语言指令（如「把正文设为小四号仿宋，首行缩进2字符」）解析为结构化配置，一次完成排版
- **Z Mode（多步骤工作流）**：通过自然语言描述多步骤工作流，如「1. 将标题2变为标题1；2. 设置多级列表；3. 再排版，正文小四」→ LLM 解析为工作流 JSON，按序执行
- **排版引擎**：批量设置字体、字号、颜色、段落格式、文本替换、标题升降级、多级列表、页面边距、页眉页脚、连续空行合并、更新域
- **预设规则**：公文标准、论文标准等常用排版预设
- **可扩展**：支持切换 LLM 后端（OpenAI、本地模型等）

## 排版逻辑

排版引擎（`WordProcessor`）对每条 `scope_rule` 执行**双重应用**：既修改 Word 文档的样式定义，又批量设置已有内容的直接格式。

### 双重应用

| 操作 | 1. 样式定义 | 2. 批量直接格式 |
|------|-------------|-----------------|
| **set_global_styles** | 修改 Normal、Heading 1 等样式的字体定义 | 遍历匹配范围的 runs，直接应用字体 |
| **paragraph_shaping** | 修改样式的段落格式（行距、缩进等） | 遍历匹配范围的段落，直接应用段落格式 |

- **样式定义**：新输入的内容将自动继承，样式面板中会显示更新后的设置
- **批量设置**：确保已有内容立即生效

### Heading 样式处理

Heading 1~9 默认使用**主题字体**（`w:asciiTheme`、`w:eastAsiaTheme` 等），在 OOXML 中会覆盖显式字体。排版引擎在设置字体时会先移除这些主题属性，使显式字体（如黑体、仿宋）正确生效。

### 作用范围与 Word 样式映射

| target_scope | 对应的 Word 样式 |
|--------------|------------------|
| body | Normal（正文） |
| heading_1 ~ heading_9 | Heading 1 ~ Heading 9 |
| heading | Heading 1 ~ 9 全部 |
| caption | Caption（题注） |
| all | Normal + Heading 1~9 + Caption |

### 执行流程

```
TypesettingIntent → to_scope_rules() → 逐条 ScopeRule
  ├─ font_config  → set_global_styles（字体、字号、颜色、加粗等）
  ├─ paragraph_config → paragraph_shaping（首行缩进、行距、段前段后、对齐）
  └─ replacements → semantic_replace（文本替换，含表格）
```

## 安装

```bash
pip install -r requirements.txt
```

## 配置

复制 `.env.example` 为 `.env`，填入以下环境变量（使用 LLM 解析意图时必需）：

```bash
cp .env.example .env
```

| 变量 | 说明 |
|------|------|
| OPENAI_API_KEY | API Key（必需） |
| OPENAI_MODEL | 模型名称，如 deepseek-chat、gpt-4o-mini（默认 deepseek-chat） |
| OPENAI_BASE_URL | 可选，API 地址，用于代理或本地服务 |

## 使用方式

### 主入口函数

```python
from src.main import process_document

# Y Mode：单次再排版
process_document(
    file_path="report.docx",
    user_prompt="正文小四仿宋，首行缩进2字符",
    preset="official",
)

# Z Mode：多步骤工作流（mode 可选 "y"、"z"、"auto"，默认 auto 自动判断）
process_document(
    file_path="report.docx",
    user_prompt="1. 将标题2变为标题1；2. 为文档设置多级列表；3. 再排版，正文小四、标题一号",
    mode="z",
)
```

### 命令行

```bash
# 从项目根目录运行
# Y Mode（单次再排版，默认 auto 会根据指令自动判断）
python -m src.main sample.docx "正文小四仿宋，首行缩进2字符"

# Z Mode（多步骤工作流，使用 -m z 强制指定）
python -m src.main sample.docx "1. 将标题2变为标题1；2. 设置多级列表；3. 再排版，正文小四" -m z

# 指定输出路径与预设
python -m src.main sample.docx "正文小四仿宋" -o output.docx -p thesis
```

### 使用 OpenAI 解析意图

```python
from llm.connector import OpenAIConnector
from src.main import process_document

connector = OpenAIConnector(model="gpt-4o-mini")
process_document(
    "report.docx",
    "把正文设为小四号仿宋，首行缩进2字符，1.5倍行距",
    preset="official",
    llm_connector=connector,
)
```

## 项目结构

```
├── src/
│   ├── main.py              # 主入口 process_document()
│   ├── config/
│   │   ├── presets.py       # Config 预设规则
│   │   └── rules_loader.py  # 排版规则 JSON 加载器
│   ├── prompts/
│   │   ├── llm_prompt.py    # Y Mode 意图解析提示词
│   │   └── workflow_prompt.py # Z Mode 工作流解析提示词
│   ├── schemas/typesetting.py  # Pydantic Schema（含 Workflow）
│   ├── intent/
│   │   ├── parser.py        # IntentParser（Y Mode）
│   │   └── workflow_parser.py # WorkflowParser（Z Mode）
│   ├── formatter/engine.py  # WordProcessor 排版引擎
│   └── dispatcher/
│       ├── command.py      # CommandDispatcher（Y Mode）
│       └── workflow.py     # WorkflowDispatcher（Z Mode）
├── docs/
│   └── TYPESETTING_RULES.md # 排版规则说明文档
├── typesetting_rules.json   # 排版规则 JSON 配置（可读可编辑）
├── llm/connector.py         # LLM 连接器
├── requirements.txt
└── .env.example
```

## 预设

| 名称 | 说明 |
|------|------|
| official | 公文标准（仿宋小四、首行缩进2字符、1.5倍行距） |
| thesis | 论文标准（宋体小四、首行缩进、段前段后6pt） |
| default | 默认（宋体12pt、左对齐） |

## 排版规则与 Schema

**Schema 优先**：后续开发请优先扩展 `src/schemas/typesetting.py`，再适配调用方。

- **[docs/SCHEMA.md](docs/SCHEMA.md)**：Schema 规范与扩展指南
- **[docs/SCHEMA_入门.md](docs/SCHEMA_入门.md)**：Schema 入门（小白版）
- **[docs/WORKFLOW.md](docs/WORKFLOW.md)**：Z Mode 工作流文档（步骤类型、JSON Schema、使用示例）
- **[docs/TYPESETTING_RULES.md](docs/TYPESETTING_RULES.md)**：排版规则说明（人类可读）
- **typesetting_rules.json**：排版规则 JSON 配置（与 Schema 对齐）

规则加载示例（统一返回 Schema 类型）：

```python
from src.config import load_rules_document, get_preset_as_scope_rules

# 加载规则文档（RulesDocument）
doc = load_rules_document()

# 获取预设的 ScopeRule 列表
scope_rules = get_preset_as_scope_rules("official")
```

## License

MIT
