# Claude Code GIS — AI 驱动的 ArcGIS Pro 自动化制图系统

用自然语言操控 ArcGIS Pro，从数据扫描到批量出图，全流程自动化。

> **核心思路**：不造轮子。Agent 层用最强的 Claude Code，执行层用成熟的 ArcPy。我们做的是中间的桥——把 GIS 领域的知识和模式翻译给 AI。

## 演示（文字版）

```
用户: 帮我把XX镇12个地块的宗地数据生成勘测定界图

Claude Code:
  1. 扫描 GDB → 发现"宗地"图层，12个要素
  2. 读取布局模板 → 17个元素，识别标题/比例尺/指北针/四角坐标
  3. 生成 ArcPy 脚本:
     - importDocument(模板)
     - 循环: definitionQuery → camera → fill() → exportToPDF
  4. 执行 → 12张PDF输出完成
  5. health_check → 无文字溢出，所有元素在页面内 ✓

耗时: ~2分钟（含生成和验证）
```

## 和 Esri 官方 Pro Assistant 的定位

Esri 在 ArcGIS Pro 3.6 中推出了 AI Assistant（Beta），支持 Help/Perform Actions/ArcPy Code Generation 等模式。

本系统与它不是替代关系，而是互补：

- **Pro Assistant** 是一个通用 AI 面板，适合单次操作和代码生成
- **本系统** 是一个专门优化批量制图流程的 Claude Code Skill，专注于"扫描→模板→出图"的完整循环

两者在各自擅长的场景下使用。如果日常任务可以通过 Pro Assistant 完成，就用它。

## 项目结构

```
claude-code-gis/
├── skills/arcpy-workflow/    ← Claude Code Skill（核心）
│   └── SKILL.md              ← AI 操作手册：API速查+模式+坑
├── lib/
│   └── pagx_layout.py        ← DOM风格布局编辑器
├── examples/
│   └── demo_workflow.py      ← 批量制图工作流示例
├── LICENSE
└── README.md
```

## 快速开始

### 1. 安装 Claude Code

```bash
npm install -g @anthropic-ai/claude-code
```

### 2. 安装本 Skill

```bash
# 克隆项目
git clone https://github.com/<your-username>/claude-code-gis.git

# 安装 Skill 到 Claude Code
mkdir -p ~/.claude/skills
cp -r skills/arcpy-workflow ~/.claude/skills/
```

### 3. 启动 Claude Code

```bash
claude
```

然后直接说你的需求："帮我把这个 GDB 里的宗地数据生成勘测定界图。"

Claude Code 会自动加载 arcpy-workflow Skill，生成并执行 ArcPy 脚本。

### 4. 环境要求

- **ArcGIS Pro** 3.0+（执行脚本）
- **ArcPy** 可用（ArcInfo 许可）
- **Python** 3.11+（Claude Code 运行环境）

## 核心能力

### 自然语言 → GIS 操作

| 你说 | 系统做 |
|------|--------|
| "把这些 shp 转成 CGCS2000" | 批量投影转换 |
| "生成这个镇的勘测定界图" | 扫描数据→加载模板→逐宗地出图 |
| "检查图层有没有拓扑错误" | 空间分析+生成质检报告 |
| "把这个布局的标题改成红色16号字" | PagxLayout DOM操作 |

### PagxLayout — 像改网页一样改布局

```python
from pagx_layout import PagxLayout
layout = PagxLayout.from_arcpy(ly)

# CSS 选择器找元素
el = layout.find_one("[name='标题']")
el.font_size = 16     # 字号
el.color = "#DC2828"  # 颜色
el.bold = True        # 加粗

# 批量填充
layout.fill({
    "标题": "XX镇XX村一组勘测定界图",
    "地块号": "宗地001",
    "面积": "123.45亩",
})

# 导出前体检
layout.health_check()
# → [文字溢出] '测绘单位名称' 框太窄
# → [对比度过低] '坐标标注' 颜色太白看不清

layout.map_series.enabled = False
ly.exportToPDF("output.pdf", resolution=200)
```

传统 ArcPy 方式需要 7 层 CIM 嵌套才能改一个颜色，PagxLayout 一句话搞定。

### 技能库

Skill 文件包含 AI 需要的全部上下文：
- **GIS 操作速查**：数据描述、游标、投影、裁剪、缓冲、选择
- **批量制图完整流程**：importDocument → MapSeries禁用 → 循环出图
- **坐标系 WKID 表**：CGCS2000/WGS84/Beijing54/Xian80
- **9 条常见坑**：中文路径/CIM版本/labelPoint/PermissionError/...
- **坐标转换公式**：地图坐标→页面坐标

## 设计理念

1. **不重复造 Agent** — Claude Code 已经是顶级 Agent，我们专注做领域适配
2. **脚本生成而非 API 封装** — 生成和执行纯文本 Python 脚本，对 AI 最友好
3. **经验即知识库** — Skill 文件沉淀的是项目实战中踩出来的坑和模式
4. **渐进式** — 不需要改现有工作流，Piece by piece 替代手动操作

## 适用场景

- 勘测定界图批量生产
- 用地审批审查图生成
- 林草资源调查图件输出
- 国土空间规划图件制作
- 任何需要从数据→布局→PDF 的 GIS 出图流程

## 许可证

MIT

## 作者

一个每天都在用 GIS 干活的人，用 AI 让自己少加班。
