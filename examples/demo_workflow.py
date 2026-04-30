# -*- coding: utf-8 -*-
"""
批量勘测定界图生产 — 完整工作流示例。

需求：从 File Geodatabase 的"宗地"图层中，
      为每个要素生成一张勘测定界图 PDF。
"""

import arcpy.mp as mp
import os
import shutil
import sys

# 如果使用 PagxLayout，取消下面注释并修改路径
# sys.path.insert(0, r"<项目目录>/lib")
# from pagx_layout import PagxLayout

# ── 配置 ──────────────────────────────────────────────────────
TEMPLATE_PAGX = r"your_project/templates/勘测定界图模板.pagx"
BLANK_APRX = r"your_project/Blank.aprx"
GDB_PATH = r"your_project/data/宗地数据.gdb"
FEATURE_CLASS = "宗地"
OUTPUT_DIR = r"your_project/output"
# ────────────────────────────────────────────────────────────────

arcpy.env.overwriteOutput = True
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 1. 从空白模板创建项目
shutil.copy2(BLANK_APRX, os.path.join(OUTPUT_DIR, "项目.aprx"))
aprx = mp.ArcGISProject(os.path.join(OUTPUT_DIR, "项目.aprx"))
layout = aprx.importDocument(TEMPLATE_PAGX)

# 2. 禁用 MapSeries（脚本手动控制）
cim_layout = layout.getDefinition('V3')
cim_layout.mapSeries = None
layout.setDefinition(cim_layout)

# 3. 获取地图框和要素图层
mf = layout.listElements("MAPFRAME_ELEMENT")[0]
m = mf.map
layer = m.listLayers(FEATURE_CLASS)[0]

# 4. 读取宗地数据
oids = []
with arcpy.da.SearchCursor(os.path.join(GDB_PATH, FEATURE_CLASS),
                           ["OBJECTID", "ZDDM", "SHAPE@"]) as cursor:
    for row in cursor:
        oids.append({"oid": row[0], "code": row[1], "shape": row[2]})

# 5. 批量导出
for i, feat in enumerate(oids):
    # 过滤当前宗地
    layer.definitionQuery = f"OBJECTID = {feat['oid']}"

    # 定位到宗地中心
    lp = feat["shape"].labelPoint
    mf.camera.X = lp.X
    mf.camera.Y = lp.Y
    mf.camera.scale = 500  # 1:500

    # 改文字
    for el in layout.listElements("TEXT_ELEMENT"):
        if el.name == "标题":
            el.text = f"{feat['code']}勘测定界图"
        elif el.name == "地块号":
            el.text = feat["code"]

    # 导出 PDF
    pdf_path = os.path.join(OUTPUT_DIR, f"{feat['code']}.pdf")
    layout.exportToPDF(pdf_path, resolution=200)
    print(f"[{i+1}/{len(oids)}] {feat['code']} → {pdf_path}")

    # 清理标注元素（如果循环中创建了图形元素）
    for el in list(layout.listElements("GRAPHIC_ELEMENT")):
        if el.name not in ("模板元素1", "模板元素2"):  # 白名单
            try:
                layout.deleteElement(el)
            except Exception:
                pass

aprx.save()
print(f"\n完成。共生成 {len(oids)} 张 PDF。")
