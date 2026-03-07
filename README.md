# 智能 Word 排版 Agent

基于 python-docx、LLM 和 Pydantic 的智能 Word 文档排版工具，支持自然语言指令驱动排版。

## 功能特性

- **意图解析**：将自然语言指令（如「把正文设为小四号仿宋，首行缩进2字符」）解析为结构化配置
- **排版引擎**：批量设置字体、字号、颜色、段落格式、文本替换
- **预设规则**：公文标准、论文标准等常用排版预设
- **可扩展**：支持切换 LLM 后端（OpenAI、本地模型等）

## 安装

```bash
pip install -r requirements.txt
```

## 配置

复制 `.env.example` 为 `.env`，填入 `OPENAI_API_KEY`（使用 LLM 解析意图时必需）：

```bash
cp .env.example .env
```

## 使用方式

### 主入口函数

```python
from src.main import process_document

# 处理文档，应用用户指令与预设
process_document(
    file_path="report.docx",
    user_prompt="正文小四仿宋，首行缩进2字符",
    preset="official",
)
```

### 命令行

```bash
# 从项目根目录运行
python -m src.main sample.docx "正文小四仿宋，首行缩进2字符"
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
│   │   └── llm_prompt.py    # LLM 提示词统一入口
│   ├── schemas/typesetting.py  # Pydantic Schema
│   ├── intent/parser.py     # IntentParser 意图解析器
│   ├── formatter/engine.py  # WordProcessor 排版引擎
│   └── dispatcher/command.py # CommandDispatcher 调度器
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
