# Word 文档物理实验报告处理工具集

> 🤖 本项目由 [Claude Code](https://claude.ai/code) Agent + [CC Switch](https://github.com/farion1231/cc-switch)（接入 [DeepSeek V4](https://api-docs.deepseek.com/zh-cn/) 和 [SiliconFlow](https://cloud.siliconflow.cn/me/models)）辅助完成

一套用于处理大学物理实验报告 `.docx` 文档的 Python 工具集，提供三大核心功能：

| # | 功能 | 方式 | 用途 |
|---|------|------|------|
| 1 | 🔤 格式标准化 | Python 脚本 | 以参考文档为模板，统一目标文档的字体、字号、标题样式、编号格式 |
| 2 | 📐 LaTeX → OMML 公式转换 | `convert_latex_in_docx.py` | 将 `$...$` / `$$...$$` 的 LaTeX 源码转为 Word 可渲染的公式 |
| 3 | 📊 数据提取与绘图 | Python 脚本 | 从实验报告 + 数据记录中提取表格数据，生成带拟合线的图像 |

> ⚠️ **物理实验专用** — 样式映射和数据处理逻辑均基于物理实验报告场景设计。

## 🔧 Claude Code 技能（Skills）

本项目两大核心功能封装为 Claude Code 技能，在 Claude Code 中可直接调用：

| 技能 | 功能 | 调用方式 |
|------|------|---------|
| **`docx-format-standardizer`** | 🔤 格式标准化 | 在 Claude Code 中说"统一格式"或"参考文档格式"即可触发 |
| **`experiment-data-plotter`** | 📊 数据提取与绘图 | 在 Claude Code 中提供实验报告 + 数据记录即可触发 |

技能文件位于 `.claude/skills/` 目录，可在 Claude Code 中自动加载使用。

---

## 快速开始

### 安装依赖

```bash
pip install python-docx lxml latex2mathml mathml2omml matplotlib numpy
```

### 项目结构

```
.
├── convert_latex_in_docx.py      # LaTeX → OMML 公式转换脚本
├── 电子电荷的测量实验报告.docx     # 示例实验报告（已匿名化处理）
├── CLAUDE.md                     # Claude Code 项目文档与 SOP
└── .claude/skills/               # Claude Code 技能定义
    ├── docx-format-standardizer/ # 格式标准化技能（SKILL.md + reference-script.py）
    └── experiment-data-plotter/  # 数据提取与绘图技能（SKILL.md）
```

---

## 功能一：LaTeX → OMML 公式转换

将 Word 文档中 `$...$`（行内）和 `$$...$$`（独立行）的 LaTeX 源码转换为 Word 原生 OMML 公式。

### 使用方法

```bash
python convert_latex_in_docx.py 输入.docx                   # 输出: 输入_公式版.docx
python convert_latex_in_docx.py 输入.docx 输出.docx         # 指定输出路径
```

转换后的公式在 Word / WPS 中可正常渲染和编辑。

### 注意事项

- **`latex2mathml`**：`converter.convert()` 的 `display` 参数传字符串 `"inline"` 或 `"block"`，不是布尔值
- **`mathml2omml`**：该库存在一个 bug，`__init__.py` 中 `MUnder.to_str()` 和 `MOver.to_str()` 里的 `</m:groupChr>` 应为 `</m:groupChrPr>`，每次重装包后需手动修复
- **中文 `\text{}`**：含中文的 `\text{}` 转换可能异常，会跳过不转
- **输入输出**：输出文件不能和输入文件同名
- **兼容性**：Google Docs 不支持 OMML，请使用 Word / WPS 打开

---

## 功能二：格式标准化

以参考实验报告为模板，统一目标文档的格式：

- **字体统一**：标题用黑体（SimHei），英文/数字用 Times New Roman
- **字号层级**：封面标题 36pt → 一级标题 18pt → 二级标题 15pt → 正文 12pt
- **编号清理**：移除标题中不规范的章节编号（如 `2.1`、`5.3.1`）
- **特殊区域**：自动跳过装订线区域、封面信息区域；思考题区域单独处理

在 Claude Code 中可通过 `docx-format-standardizer` 技能自动执行，或参考 `.claude/skills/docx-format-standardizer/reference-script.py` 编写自定义脚本。

---

## 功能三：数据提取与绘图

从实验报告和数据记录两个 `.docx` 文件中提取表格数据，生成格式化图像：

- 散点图 + 线性拟合线（含拟合方程和 R²）
- 异常点标注与剔除
- 关键参数标注（灵敏度、温度系数、B 常数等）
- 中文支持（SimHei 字体）
- 图片输出为 `Fig{编号}_{描述}.png`，dpi=200

在 Claude Code 中可通过 `experiment-data-plotter` 技能自动执行。

---

## 完整工作流

### Step 1：AI 生成实验报告初稿

在 Claude Code、Cherry Studio 等 AI 客户端中使用以下提示词，并以 **附件形式** 上传 **具体的实验数据**。实验指导书、参考资料等 PDF 文件可通过客户端的 **RAG（检索增强生成）** 功能自动解析，无需手动摘录内容：

> 你好，我是一名物理系的学生，请帮我完成一份实验报告，要求参考附件，以报告模板为基准，实验数据表单为数据来源，其余附件为了解实验提供参考
>
> **完整性与规范性**
> 1. 报告内容完整，需包含实验目的、实验原理、实验仪器、实验内容、数据处理、分析讨论、实验结论、思考题、参考文献等；
> 2. 报告格式规范；表格、公式与文字不直接使用贴图；
> 3. 图表清晰、规范；有横纵坐标标度与单位；
> 4. 有效数字规范；单位完整；实验结果表达规范；
> 5. 文字表达流畅，逻辑层次清晰；
>
> **正确性**
> 1. 实验数据真实可靠；
> 2. 数据处理过程完整且计算结果正确（包括公式代入、图表数据的分析处理等）；
> 3. 不确定度估算或误差分析合理且过程完整；
> 4. 图表绘制正确，拟合参数完整正确；
> 5. 思考题回答正确；
>
> **创新性**
> 能对实验现象进行深入分析；在误差分析或数据处理方法上有独到见解；对实验改进提出有价值的建议；完成实验拓展内容设计方案等；
>
> **在分析讨论、结论中**（基于实验现象及数据处理的结果进行分析讨论，给出实验结果），可以（此项并不强求每一点都做）
> 1. 分析实验结果与预期是否一致，与其它的研究结果异同。
> 2. 分析实验中异常现象成因。
> 3. 分析误差的来源，以及实验技术或方法的可能缺失。
> 4. 讨论提高或改进实验测量准确性的方法。
> 5. 讨论实验心得或建议。
> 6. 归纳一般性、概括性的概念、原则或理论的简明总结。
>
> **注意：不能篡改和编造数据。不限篇幅**

### Step 2：导出为 Word 文档

将 AI 的回答导出为 `.docx` 文件，保存到本项目根目录。

### Step 3：格式标准化三部曲（可通过 Agent 自动完成）

对导出的 Word 文档依次执行以下操作（在 Claude Code 中可交由 **Agent** 自动完成）：

**① LaTeX → OMML 公式转换**
将文档中 `$...$` / `$$...$$` 的 LaTeX 源码转为 Word 可渲染的公式：
```bash
python convert_latex_in_docx.py 报告.docx
```

**② 格式标准化**
统一字体字号、清理不规范编号（参考 `docx-format-standardizer` 技能或 `reference-script.py`）。

**③ 数据提取与绘图**
从报告和数据记录中提取表格数据，生成带拟合线的图像。

---

## 隐私说明

本仓库中的所有文档均经过匿名化处理，不包含真实姓名、学号、日期等个人信息。如需在公开仓库中使用自己的实验报告，请确保提前抹除隐私信息。

---

## 许可

仅供学习交流使用。
