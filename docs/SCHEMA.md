# 排版 Schema 规范

**Schema 优先**：后续开发请优先扩展 `src/schemas/typesetting.py`，再适配调用方。

> 新手入门：若对 Schema 概念不熟，可先看 [SCHEMA_入门.md](SCHEMA_入门.md)（小白版）。

## 模块结构

```
src/schemas/typesetting.py   # 单一数据源，所有类型定义
├── TargetScope              # 作用范围字面量
├── FontConfig               # 字体配置
├── ParagraphConfig          # 段落配置
├── ScopeRule               # 单 scope 规则（支持 font/paragraph 别名）
├── ReplaceRule             # 文本替换规则
├── TypesettingIntent       # 排版意图（LLM 输出 / 预设）
├── RuleScopeEntry          # JSON 单 scope 条目
├── PresetRules             # 预设规则集
└── RulesDocument           # typesetting_rules.json 完整 Schema
```

## 核心类型

### TargetScope

作用范围枚举，与 `typesetting_rules.json` 的 `scope_mapping` 一致：

- `body`：正文
- `heading`：全部标题
- `heading_1` ~ `heading_9`：指定级标题
- `caption`：小字题注
- `emphasis`：强调（run 级）
- `all`：全文

### ScopeRule

单作用范围的排版规则，支持 JSON 别名：

| Schema 字段        | JSON 别名 | 说明     |
|--------------------|-----------|----------|
| font_config        | font      | 字体配置 |
| paragraph_config   | paragraph | 段落配置 |

### TypesettingIntent

排版意图，支持两种模式：

1. **scope_rules**（主）：多 scope 配置列表
2. **legacy**：`global_styles` + `paragraph_config` + `target_scope`

统一入口：`intent.to_scope_rules()` 返回 `List[ScopeRule]`，供 Dispatcher 使用。

### RulesDocument

对应 `typesetting_rules.json` 的完整结构，包含：

- `font_size_reference`：字号对照表
- `indent_reference`：首行缩进对照
- `presets`：预设名称 → PresetRules
- `scope_mapping`：scope 标识 → Word 样式描述

## 函数接口规范

| 函数                     | 返回类型        | 说明                     |
|--------------------------|-----------------|--------------------------|
| `load_rules_document()`   | RulesDocument   | 加载规则文档（主接口）   |
| `get_preset_rules()`      | PresetRules     | 获取预设规则             |
| `get_scope_rule()`        | ScopeRule       | 获取单 scope 规则        |
| `get_preset_as_scope_rules()` | List[ScopeRule] | 预设转 ScopeRule 列表    |
| `format_rules_for_prompt()` | str            | 格式化为 LLM prompt 文本 |

## JSON 与 Schema 对齐

`typesetting_rules.json` 的 `presets.<name>.rules.<scope>` 结构：

```json
{
  "heading_1": {
    "font": { "name": "黑体", "size_pt": 22 },
    "paragraph": { "alignment": "center", "space_before": 12 }
  }
}
```

对应 `RuleScopeEntry`，可被 `RulesDocument.model_validate()` 解析。

## 扩展指南

1. **新增 scope**：在 `TargetScope` 和 `TARGET_SCOPES` 中添加
2. **新增配置字段**：在 `FontConfig` / `ParagraphConfig` 中添加
3. **新增规则类型**：在 `TypesettingIntent` 或新建 Schema 中添加
4. **JSON 结构变更**：同步更新 `RulesDocument`、`PresetRules`、`RuleScopeEntry`
