# Word 文档物理实验报告处理项目

本项目围绕大学物理实验报告，提供三大核心功能：

| # | 功能 | 触发方式 | 用途 |
|---|------|---------|------|
| 1 | 🔤 格式标准化 | `docx-format-standardizer` 技能 | 参考文档刷格式、统一字体字号编号 |
| 2 | 📐 LaTeX→OMML 公式转换 | `python convert_latex_in_docx.py` | LaTeX源码转Word可渲染公式 |
| 3 | 📊 数据提取与绘图 | `experiment-data-plotter` 技能 | 从报告+数据记录生成带拟合的图像 |

> **⚠️ 三个功能均为物理实验专用，勿套用到其他类型文档**

## 相关技能与记忆
- 格式标准化技能：`docx-format-standardizer`（物理实验专用，勿套用到其他类型文档）
- 实验数据提取与绘图技能：`experiment-data-plotter`（**物理实验专用，勿套用到其他类型文档** — 从实验报告+数据记录中提取数据生成图像）
- 项目技能映射：详见 `.claude/skills/experiment-data-plotter/.mapping`
- 项目记忆：详见 `memory/` 目录

# Word 文档 LaTeX 公式转换 SOP

## 核心需求
将 Word 文档中 `$...$`（行内）和 `$$...$$`（独立行）的 LaTeX 源码转换为 Word 可渲染的 OMML 公式。转换失败的公式保留原文不转换。

## 工具链
- `python-docx` / `lxml` — 读写 .docx
- `latex2mathml` — LaTeX → MathML
- `mathml2omml` — MathML → OMML

## 脚本
项目根目录有 `convert_latex_in_docx.py`：
```bash
python convert_latex_in_docx.py 输入.docx                  # 输出: 输入_公式版.docx
python convert_latex_in_docx.py 输入.docx 输出.docx        # 指定输出路径
```
处理所有段落和页眉页脚中的 `$...$` / `$$...$$`。

## 已知坑点（必读）

### latex2mathml
- `converter.convert(tex, display="inline")` — display 传**字符串** `"inline"`/`"block"`，不是布尔值 `True`/`False`

### mathml2omml（有 bug 需手动修复）
- **每次重装包后都要重新修！** 库文件 `__init__.py` 中两处 `<m:groupChrPr>` 被错误闭合为 `</m:groupChr>`（少写 `Pr`）
- 具体位置在 `MUnder.to_str()` 和 `MOver.to_str()` 方法
- 需将两处 `'</m:groupChr>'` 改为 `'</m:groupChrPr>'`
- 其 `convert()` 输出缺 `xmlns:m` 声明，脚本中用 `.replace()` 补上

### XML / lxml
- `xml:space` 属性用 `XML = '{http://www.w3.org/XML/1998/namespace}'` + `t.set(XML + 'space', 'preserve')`
- 段落清空重建会丢失 `w:pPr`，需提前保存再重新插入
- 用 `etree.register_namespace('m', NSM)` 确保序列化时带上命名空间

### Python 字符串
- LaTeX 用 raw string `r'\frac{1}{2}'`，否则 `\f` 被解析为换页符
- 中文 `\text{...}` 转换可能异常，跳过不转

### Word 文档限制
- 输出文件和输入文件不能是同一个路径，否则 PermissionError
- 转换后的公式在 Word/WPS 中可正常渲染，但 Google Docs 可能不兼容 OMML

## 命名空间
- OMML: `http://schemas.openxmlformats.org/officeDocument/2006/math`（前缀 `m:`）
- Word: `http://schemas.openxmlformats.org/wordprocessingml/2006/main`（前缀 `w:`）
