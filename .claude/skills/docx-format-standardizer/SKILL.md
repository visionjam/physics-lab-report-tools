---
name: docx-format-standardizer
description: >
  【物理实验专用】大学物理实验报告 .docx 格式标准化工具。
  按参考文档统一字体字号、去掉不规范多级编号（如 2.1, 5.3.1）。
  专用场景：物理实验报告、"大学物理实验"课程报告。
  非物理实验类文档请勿套用。Triggers on:
  "统一格式", "格式调整", "参考文档格式", "docx格式", "Word格式标准化",
  "修改标题格式", "调整字体大小", "去掉编号", "不合规范的编号",
  "格式化版", "字体统一", "标题正文格式一致", "物理实验", "实验报告".
  Keywords: docx, Word, python-docx, formatting, 物理实验.
---

# DOCX Format Standardizer（物理实验报告专用）

> ⚠️ **本技能为大学物理实验报告专用。** 样式映射基于物理实验报告模板（黑体标题 + Times New Roman 副标题 + 装订线 + 思考题区域），其他类型文档请勿直接套用。

## Purpose
按参考文档统一物理实验报告 .docx 的格式（字体、字号、标题层级），去掉不规范的多级编号。专处理大学物理实验报告。

## When to Use
- 用户提供一份参考报告和一份待处理报告，要求"格式统一"
- 需要去掉标题中 "2.1"、"5.3.1" 等不规范的章节编号
- 需要统一正文 12pt、标题 18pt、二级标题 15pt 的层级
- **仅限物理实验报告场景**

## Workflow

### Step 1: Analyze the Reference Document
First, extract formatting info from the reference document to understand the target style:

```python
from docx import Document

ref = Document("reference.docx")
for i, p in enumerate(ref.paragraphs):
    text = p.text.strip()
    if not text: continue
    font_info = ""
    if p.runs:
        r = p.runs[0]
        font_info = f"font={r.font.name}, size={r.font.size}, bold={r.font.bold}"
    print(f"P[{i}] style={p.style.name} | {font_info}")
    print(f"     text: {text[:80]}")
```

### Step 2: Establish Formatting Map
Map the reference document's hierarchy to concrete EMU values:

| Level | Reference Style | Font | Size | Bold |
|-------|----------------|------|------|------|
| 封面标题 | 实验名称（居中） | 黑体 | 36pt (457200 EMU) | Yes |
| 一级标题 | 实验目的/原理等 | 黑体 | 18pt (228600 EMU) | Yes |
| 二级标题 | Heading 3 样式 | Times New Roman | 15pt (190500 EMU) | Yes |
| 三级标题 | Heading 4 样式 | Times New Roman | 15pt (190500 EMU) | Yes |
| 正文 | Normal 样式 | 继承 | 12pt (152400 EMU) | 按需 |
| 思考题题目 | Normal + "数字."开头 | 黑体 | 15pt (190500 EMU) | Yes |

Use `docx.shared.Pt()` or direct EMU values:
```python
from docx.shared import Pt, Emu
PT12 = 152400   # 12pt
PT15 = 190500   # 15pt
PT18 = 228600   # 18pt
PT36 = 457200   # 36pt
```

### Step 3: Write the Processing Script

**Key patterns:**

**Remove non-standard heading numbering:**
```python
import re
HEADING_NUM_PATTERN = re.compile(r'^\d+(?:\.\d+)+\s+')

def remove_numbering(text):
    return HEADING_NUM_PATTERN.sub('', text)
```

**Set run font (with Chinese font support):**
```python
from docx.oxml.ns import qn

def set_run_font(run, font_name=None, font_size=None, bold=None):
    if font_name is not None:
        run.font.name = font_name
        rpr = run._element.get_or_add_rPr()
        rFonts = rpr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = run._element.makeelement(qn('w:rFonts'), {})
            rpr.insert(0, rFonts)
        rFonts.set(qn('w:eastAsia'), font_name)
    if font_size is not None:
        run.font.size = Emu(font_size)
    if bold is not None:
        run.font.bold = bold
```

**Detect question paragraphs (keep numbering for 思考题):**
```python
QUESTION_PATTERN = re.compile(r'^\d+\.\s+')
def is_question_paragraph(text):
    return bool(QUESTION_PATTERN.match(text))
```

### Step 4: Handle Special Zones

**装订线区域**: Look for paragraphs with single characters like `装`, `订`, `线`, `┊` in CENTER alignment — skip them.

**封面信息区域**: The cover page (大学名称, 实验报告, 学生信息 rows) typically already has correct formatting (18pt). Skip if already sized.

**思考题区域**: 
- Questions ("1. xxx?", "2. xxx?"): 黑体 15pt bold
- "答：" lines: 12pt body
- Answer content: 12pt body

### Step 5: Process Tables
If the document has tables, iterate table cells and apply body text formatting:

```python
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                set_all_runs_font(p, None, PT12, None)
```

### Full Processing Script Template

See [reference-script.py](reference-script.py) for a complete working example (the script used to process 传感器特性研究实验报告.docx).

## Common EMU Values Reference
- 8pt = 101600, 9pt = 114300, 10pt = 127000
- 10.5pt = 133350 (Normal default), 12pt = 152400
- 14pt = 177800, 15pt = 190500, 16pt = 203200
- 18pt = 228600, 22pt = 279400, 36pt = 457200

## Common Fonts for Chinese Docs
- 黑体 (SimHei) — headings
- 宋体 (SimSun) — body (Chinese)
- Times New Roman — body/subsections (Western)
- 楷体 (KaiTi) — occasional emphasis
