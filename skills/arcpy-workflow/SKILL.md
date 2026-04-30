---
name: arcpy-workflow
description: ArcPy GIS 智能化工作流。当用户需要进行 GIS 数据处理、空间分析、自动化制图、批量数据操作时触发。支持自然语言描述需求，自动生成并执行 ArcPy 脚本。
---

# ArcPy 智能化工作流

## 执行模式

生成脚本 → 写临时文件 → ArcGIS Pro Python 执行 → 解析 JSON 输出。

```bash
"<ArcGIS Pro Python路径>" <临时脚本路径>
# 例如: "C:/Program Files/ArcGIS/Pro/bin/Python/envs/arcgispro-py3/python.exe" %TEMP%/arcpy_script.py
```

脚本模板：

```python
import json, sys, os, traceback
try:
    import arcpy
    arcpy.env.overwriteOutput = True

    # === 用户代码 ===

    print(json.dumps({"success": True, "data": ...}, ensure_ascii=False))
except Exception as e:
    traceback.print_exc()
    print(json.dumps({"success": False, "error": str(e)}, ensure_ascii=False))
```

## 环境

- **Python**: `C:/Program Files/ArcGIS/Pro/bin/Python/envs/arcgispro-py3/python.exe`
- **ArcPy**: 3.4, ArcInfo
- **路径**: 统一 `r"..."` 原始字符串，输出前 `os.makedirs(dir, exist_ok=True)`
- **临时脚本**: `%TEMP%/arcpy_script.py`（Windows 临时目录）

## 常用 API 速查

### 数据操作

```python
# 描述数据
desc = arcpy.Describe(r"D:/data.shp")
desc.name, desc.dataType, desc.shapeType  # 等

# 列出数据
arcpy.env.workspace = r"D:/geodatabase.gdb"
arcpy.ListFeatureClasses()
arcpy.ListTables()
arcpy.ListRasters()

# 字段
arcpy.ListFields(r"D:/data.shp")  # → [name, type, length]

# 游标
with arcpy.da.SearchCursor(r"D:/data.shp", ["FIELD1", "SHAPE@"]) as cur:
    for row in cur: ...

with arcpy.da.UpdateCursor(r"D:/data.shp", ["FIELD"]) as cur:
    for row in cur:
        row[0] = new_val
        cur.updateRow(row)
```

### 空间分析

```python
arcpy.Project_management(in_fc, out_fc, arcpy.SpatialReference(4490))
arcpy.Clip_analysis(in_fc, clip_fc, out_fc)
arcpy.Buffer_analysis(in_fc, out_fc, "500 meters")
arcpy.Select_analysis(in_fc, out_fc, where_clause="FIELD > 10")
# 空间选择
arcpy.Select_analysis(in_fc, out_fc, select_features=boundary_fc, overlap_type="WITHIN")
```

### 坐标系 WKID

| 坐标系 | WKID |
|--------|------|
| WGS 84 | 4326 |
| CGCS2000 地理 | 4490 |
| CGCS2000 3-degree GK CM 105E | 4525 |
| CGCS2000 3-degree GK CM 120E | 4527 |
| Web Mercator | 3857 |
| Beijing 1954 | 4214 |
| Xian 1980 | 4610 |

---

## 批量制图完整流程

**核心模式** — 最常用的场景（勘测定界图、审查图）：

```python
import arcpy.mp as mp, os, shutil

# 1. 从空白模板创建项目（路径改为你的实际路径）
shutil.copy2(r"your_project/Blank.aprx", r"your_project/output/项目.aprx")
aprx = mp.ArcGISProject(r"your_project/output/项目.aprx")
layout = aprx.importDocument(r"your_project/templates/布局.pagx")

# 2. 禁用 MapSeries（通过 CIM）
cim_layout = layout.getDefinition('V3')
cim_layout.mapSeries = None
layout.setDefinition(cim_layout)

# 3. 获取地图框和图层
mf = layout.listElements("MAPFRAME_ELEMENT")[0]
m = mf.map
layer = m.listLayers("目标图层名")[0]

# 4. 批量导出
for oid, name, ... in feature_list:
    # 过滤 + 定位
    layer.definitionQuery = f"OBJECTID = {oid}"
    mf.camera.scale = 500  # 1:500

    # 改文字
    for el in layout.listElements("TEXT_ELEMENT"):
        if el.name == "标题":
            el.text = f"{name}勘测定界图"

    # 导出
    layout.exportToPDF(rf"your_project/output/{name}.pdf", resolution=200)

aprx.save()
```

---

## 布局元素操作（PagxLayout）

**禁止使用原始 CIM 路径修改元素样式，必须用 PagxLayout。**
库: `lib/pagx_layout.py`

```python
import sys
sys.path.insert(0, r"<项目目录>/lib")
from pagx_layout import PagxLayout

layout = PagxLayout.from_arcpy(ly)

# 查元素
el = layout.find_one("[name='标题']")
for e in layout.find("[name*='文本']"): ...  # 批量

# 改属性
el.font_size = 16          # 字号
el.color = "#DC2828"       # 颜色(HEX)
el.bold = True             # 加粗
el.font_family = "SimHei"  # 字体
el.h_align = "Center"      # Left/Center/Right
el.text = "新文字"

# 批量填充
layout.fill({
    "标题": "XX镇石渔村一组勘测定界图",
    "地块号": "宗地001",
    "面积": "123.45亩",
})

# 地图框
mf = layout.find_one("[type='CIMMapFrame']")
mf.set_position(30, 30, 160, 240)  # x,y,w,h(mm)

# 页面
layout.page.set_size(420, 297)   # A3
layout.map_series.enabled = False

# 调试
layout.tree()             # 打印元素树
layout.health_check()     # 体检：字号/颜色/溢出/重叠 问题

# 修改实时生效，直接导出
ly.exportToPDF(r"your_project/output.pdf", resolution=200)
```

选择器: `[name='x']`, `[name*='x']`, `[type='CIMMapFrame']`, 可组合
类型: CIMMapFrame, CIMGraphicElement, CIMGroupElement, CIMMarkerNorthArrow

---

## 常见坑

| 坑 | 解决 |
|---|---|
| 中文路径 OSError | 用 ASCII 路径 |
| CIM V3_4 报错 | 必须 `getDefinition('V3')` |
| `listElements()` 遍历时删除 | 用 `list()` 包裹 |
| 多边形 extent 中心在外部 | 用 `shape.labelPoint` 代替 `extent.center` |
| `createGraphicElement` ValueError | name 必须用关键字 `name="xxx"` |
| 字段值含换行符 | `str(val).strip().replace("\n","").replace("\r","")` |
| PDF 目标被占用 | PermissionError → 跳过或 `_v2` 后缀 |
| MapSeries 脚本模式不稳定 | 手动 camera 控制更可靠 |
| 项目目录数据丢失 | 背景数据放模板同目录专用 GDB |

## 坐标转换：地图坐标 → 页面坐标

```python
# 地图框在页面上的位置(mm)
mf_x, mf_y = mf.elementPositionX, mf.elementPositionY
mf_w, mf_h = mf.elementWidth, mf.elementHeight

# 地图框可见范围(地图单位)
sr = m.spatialReference
mpu = sr.metersPerUnit if sr.type == "Projected" else 1.0
vis_w = mf.camera.scale * mf_w * 0.0393701 / 39.3701 / mpu
vis_h = mf.camera.scale * mf_h * 0.0393701 / 39.3701 / mpu

def map_to_page(mx, my):
    """地图坐标 → 页面坐标(mm)，用于创建引线/标注"""
    left = mf.camera.X - vis_w / 2
    bottom = mf.camera.Y - vis_h / 2
    px = mf_x + (mx - left) / vis_w * mf_w
    py = mf_y + (my - bottom) / vis_h * mf_h
    return (px, py)
```

## 图层连接属性更新

```python
for lyr in m.listLayers():
    if lyr.name == "目标图层":
        lyr.updateConnectionProperties(
            lyr.connectionProperties,
            {'dataset': '要素类名',
             'workspace_factory': 'File Geodatabase',
             'connection_info': {'database': r'GDB路径'}}
        )
```

## CIM 颜色修改（仅图形元素，文字推荐 PagxLayout）

```python
cim = element.getDefinition('V3')
# 结构: CIMGraphicElement → graphic → symbol → symbol(.symbol) → symbolLayers
# 边框 CIMSolidStroke: .color.values = [R,G,B,Alpha]
# 填充 CIMSolidFill: .color.values = [R,G,B,Alpha]
# Alpha=100 不透明
element.setDefinition(cim)
```
