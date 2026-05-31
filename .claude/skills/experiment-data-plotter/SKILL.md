---
name: experiment-data-plotter
description: "Use when the user has a physics experiment report (.docx) and a corresponding data record (.docx) and needs to extract table data and generate formatted plots with linear fits, R^2, and sensitivity annotations. Trigger phrases include: '帮我绘制图像', '生成图表', '绘图', 'extract data and plot' in the context of physics experiment reports."
---

# 物理实验报告数据提取与绘图 — 物理实验专用

> **⚠️ 适用范围：仅限物理实验类文档**
> 本技能专为大学物理实验课程设计（如传感器综合实验）。
> **勿套用到其他类型文档**（如化学实验、工程报告、商业图表等）。

## 概述
从物理实验报告和数据记录两个 .docx 文件中提取表格数据，根据报告分析部分的需求生成格式化图像。

## 触发条件
- 用户同时提供物理实验报告 + 数据记录两个 .docx 文件
- 用户提到"绘图"、"生成图像"、"画图"、"plot" 等
- 数据记录中包含传感器实验数据表格（应变、电容、温度传感器等）

## 工作步骤

### 1. 读取文档
```python
import docx
report = docx.Document('实验报告.docx')
data = docx.Document('数据记录.docx')
```
- 提取所有段落文本（定位"数据记录处理"分析部分）
- 提取所有表格数据

### 2. 匹配需求
报告分析段（如"根据表1数据..."）告诉你要画什么、怎么拟合。数据记录对应表号提供原始数据。

**无 scipy 时用 numpy 实现线性回归：**
```python
def linregress(x, y):
    x, y = np.array(x, dtype=float), np.array(y, dtype=float)
    x_m, y_m = np.mean(x), np.mean(y)
    slope = np.sum((x - x_m) * (y - y_m)) / np.sum((x - x_m)**2)
    intercept = y_m - slope * x_m
    y_pred = slope * x + intercept
    r2 = 1 - np.sum((y - y_pred)**2) / np.sum((y - y_m)**2)
    return slope, intercept, r2
```

### 3. 绘图格式规范
- **字体**: `SimHei` + `axes.unicode_minus=False`
- **避免 Unicode 上标**: SimHei 无 ²/⁻/¹ glyph → 用 `R^2`、`K^-1`、`℃^-1`
- **数据点**: `'o'`/`'s'`/`'^'` 标记，不同颜色
- **拟合线**: 虚线 `'--'`，透明度 0.5-0.6
- **图例**: 含方程 `y=kx+b` + `R^2`
- **标注框**: `bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.8)`
- **异常点**: 空心标记 `markerfacecolor='none'` + 图例说明，拟合时剔除
- **输出**: `Fig{编号}_{描述}.png`，dpi=200

### 常见传感器映射 — 以传感器综合实验为例
本表基于大学物理传感器综合实验，其他物理实验可参照此模式提取对应表格数据。
| 传感器 | X | Y | 拟合 |
|--------|---|---|------|
| 应变(单臂/半桥/全桥) | 砝码质量W(g) | 输出电压U(mV) | 线性，三条曲线同图 |
| 电容(左移/右移) | 位移x(mm) | 输出电压U(mV) | 线性，两条曲线同图 |
| RTD | 温度T(℃) | 电阻R(kΩ) | 线性，剔除异常点 |
| PTC | 温度T(℃) | 电阻R(kΩ) | 曲线(显示居里点) |
| NTC lnR-1/T | 1/T(K^-1) | lnR | 线性，求B常数 |
| PN结 | 温度T(℃) | 电压U(V) | 线性，求温度系数 |
| LM35 | 温度T(℃) | 输出电压U(mV) | 线性，求灵敏度 |
| AD590 | 热力学温度T(K) | 电流I(μA) | 线性，剔除异常点 |
| 热电偶 | 温差ΔT(℃) | 温差电动势ΔU(μV) | 线性，灵敏度/100 |
| 热电堆 | 温差ΔT(℃) | 温差电动势ΔU(μV) | 线性，灵敏度/100 |

### 已知问题
- `python-docx` 表格中有时出现合并单元格，行读取时需去重
- 数据记录中的表格内容通常与报告中一致
- matplotlib 3.10+ 有 font fallback，但 SimHei 的 glyph 缺失仍然会 warn，用 ASCII 替代最保险
