# Schema 入门指南（小白版）

> 如果你觉得 `SCHEMA.md` 太专业，先看这篇。用大白话解释「排版规则」是怎么被程序理解和执行的。

---

## 一、Schema 是什么？

**一句话**：Schema 就是「格式说明书」，告诉程序：排版规则长什么样、有哪些字段、怎么填。

就像你去打印店，店员给你一张表格让你填「字体、字号、对齐方式」——Schema 就是那张表格的电子版，程序按它来解析和校验数据。

---

## 二、文档里有哪些「东西」？

可以把 Word 文档想象成一本小册子，里面有不同「角色」：

| 角色 | 英文名 | 通俗理解 |
|------|--------|----------|
| 正文 | body | 普通段落，像报纸正文 |
| 一级标题 | heading_1 | 最大的标题，比如「第一章」 |
| 二级标题 | heading_2 | 次一级，比如「1.1 节」 |
| 三级标题 | heading_3 | 更小一级 |
| … | heading_4 ~ heading_9 | 以此类推 |
| 题注 | caption | 图表下方的小字说明 |
| 全文 | all | 整篇文档 |

每个角色都可以有自己的「打扮」：用什么字体、多大字号、怎么对齐。

---

## 三、一条规则长什么样？

一条规则 = **给谁** + **怎么打扮**。

### 3.1 字体打扮（FontConfig）

| 字段 | 什么意思 | 例子 |
|------|----------|------|
| name | 字体名 | 仿宋、黑体、宋体 |
| size_pt | 字号（磅） | 12 = 小四号，22 = 二号 |
| color | 颜色 | #000000 黑色 |
| bold | 加粗 | true / false，默认各级标题加粗 |
| italic | 斜体 | true / false |
| underline | 下划线 | true / false |

### 3.2 段落打扮（ParagraphConfig）

| 字段 | 什么意思 | 例子 |
|------|----------|------|
| first_line_indent | 首行缩进多少 | 24 ≈ 缩进 2 个字符 |
| line_spacing | 行距倍数 | 1.5 = 1.5 倍行距 |
| alignment | 对齐方式 | left 左对齐，center 居中，justify 两端对齐 |
| space_before / space_after | 段前、段后留白 | 6 = 段前 6 磅 |

---

## 四、一个完整例子

用户说：「正文小四仿宋，首行缩进 2 字符；一级标题黑体二号居中」

程序会把它变成两条规则：

```json
[
  {
    "target_scope": "body",
    "font": { "name": "仿宋", "size_pt": 12 },
    "paragraph": { "first_line_indent": 24, "line_spacing": 1.5 }
  },
  {
    "target_scope": "heading_1",
    "font": { "name": "黑体", "size_pt": 22 },
    "paragraph": { "alignment": "center" }
  }
]
```

- 第一条：给 **正文** 用仿宋 12pt，首行缩进 24pt
- 第二条：给 **一级标题** 用黑体 22pt，居中

---

## 五、数据从哪来、到哪去？

```
用户说一句话
    ↓
LLM 理解并输出 JSON（按 Schema 格式）
    ↓
解析成 TypesettingIntent（排版意图）
    ↓
to_scope_rules() 转成 ScopeRule 列表
    ↓
Dispatcher 一条条应用到 Word 文档
```

- **TypesettingIntent**：可以理解为「用户想怎么排版的完整描述」
- **ScopeRule**：可以理解为「给某一个角色的一条具体打扮指令」
- **to_scope_rules()**：把完整描述拆成一条条具体指令

---

## 六、常见问题

### Q：font 和 font_config 有什么区别？

没有区别，两种写法都可以。`font` 是简写，`font_config` 是完整名，程序都认得。

### Q：字号 12、22 是什么意思？

单位是「磅」（pt）。小四号 ≈ 12pt，二号 ≈ 22pt。详见 `typesetting_rules.json` 里的 `font_size_reference`。

### Q：首行缩进 24 是什么意思？

24 磅 ≈ 2 个汉字的宽度。1 字符 ≈ 12pt。

### Q：想改「二级标题」的样式，怎么写？

把 `target_scope` 写成 `heading_2`，后面跟 `font` 和 `paragraph` 即可。

---

## 七、下一步

- 想改规则：编辑 `typesetting_rules.json`
- 想加新功能：先看 `SCHEMA.md` 的扩展指南，再改 `src/schemas/typesetting.py`
- 想理解完整技术细节：看 `SCHEMA.md`
