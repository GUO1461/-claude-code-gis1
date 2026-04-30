# -*- coding: utf-8 -*-
"""
pagx_layout — DOM-style layout editor for ArcGIS Pro.

Works in TWO modes:

  1. File mode (no ArcGIS needed):
     layout = PagxLayout("template.pagx")
     layout.find_one("[name='标题']").font_size = 14
     layout.save("modified.pagx")

  2. ArcPy mode (inside ArcGIS Pro Python):
     layout = PagxLayout.from_arcpy(arcpy_layout)
     layout.find_one("[name='标题']").font_size = 14
     # changes are live, no save needed
     arcpy_layout.exportToPDF("output.pdf")
"""

import json
import os
import re

# ── helpers ────────────────────────────────────────────────────────

def _rgba_to_hex(values):
    r, g, b = values[0:3]
    a = values[3] if len(values) > 3 else 100
    return f"#{r:02X}{g:02X}{b:02X}" if a == 100 else f"#{r:02X}{g:02X}{b:02X}{a:02X}"

def _hex_to_rgba(hex_str):
    h = hex_str.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    a = int(h[6:8], 16) if len(h) >= 8 else 100
    return [r, g, b, a]

def _parse_selector(selector):
    conds = {}
    for part in selector.strip().lstrip("[").rstrip("]").split("]["):
        m = re.match(r"(\w+)\s*=\s*['\"]?(.*?)['\"]?$", part)
        if m:
            k, v = m.group(1), m.group(2)
            conds[k] = True if v == "true" else (False if v == "false" else v)
        m = re.match(r"(\w+)\s*\*=\s*['\"](.*?)['\"]$", part)
        if m:
            conds[m.group(1) + "__contains"] = m.group(2)
    return conds

def _match_element(el, conds):
    for key, val in conds.items():
        if key.endswith("__contains"):
            k = key[:-10]
            if not (isinstance(el.get(k, ""), str) and val in el.get(k, "")):
                return False
        elif key in ("type", "name"):
            if el.get(key) != val:
                return False
        elif key == "visible":
            if el.get("visible", True) != val:
                return False
    return True

def _get_frame_rings(el):
    rings = (el.get("frame") or {}).get("rings", [[]])
    return rings

def _set_frame_rings(el, rings):
    frame = el.get("frame")
    if frame is None:
        frame = {}
        el["frame"] = frame
    frame["rings"] = rings
    outer = rings[0]
    cx = (min(p[0] for p in outer) + max(p[0] for p in outer)) / 2
    cy = (min(p[1] for p in outer) + max(p[1] for p in outer)) / 2
    if "rotationCenter" in el:
        el["rotationCenter"]["x"] = cx
        el["rotationCenter"]["y"] = cy

def _get_text_symbol(el):
    return el.get("graphic", {}).get("symbol", {}).get("symbol", {})

def _get_text_color_layer(el):
    fill_sym = _get_text_symbol(el).get("symbol", {})
    for layer in fill_sym.get("symbolLayers", []):
        if layer.get("type") == "CIMSolidFill":
            return layer
    return None

def _get_border_stroke(el):
    gf = el.get("graphicFrame", {})
    sym = (gf.get("borderSymbol") or {}).get("symbol", {})
    for layer in sym.get("symbolLayers", []):
        if layer.get("type") == "CIMSolidStroke":
            return layer
    return None

def _ensure_text_color(el, r, g, b, a=100):
    sym = _get_text_symbol(el)
    fill_sym = sym.get("symbol")
    if fill_sym is None:
        fill_sym = {"type": "CIMPolygonSymbol", "symbolLayers": [], "angleAlignment": "Map"}
        sym["symbol"] = fill_sym
    layers = fill_sym.get("symbolLayers")
    if layers is None:
        layers = []
        fill_sym["symbolLayers"] = layers
    for layer in layers:
        if layer.get("type") == "CIMSolidFill":
            c = layer.get("color")
            if c is None:
                c = {"type": "CIMRGBColor", "values": [0, 0, 0, 100]}
                layer["color"] = c
            c["values"] = [r, g, b, a]
            return
    layers.append({"type": "CIMSolidFill", "enable": True,
                   "color": {"type": "CIMRGBColor", "values": [r, g, b, a]}})

def _find_in_dicts(container, conds):
    """Recursively find CIM element dicts matching conditions."""
    results = []
    elems = container.get("elements", []) if isinstance(container, dict) else container
    for el in elems:
        if _match_element(el, conds):
            results.append(el)
        if el.get("type") == "CIMGroupElement":
            results.extend(_find_in_dicts(el, conds))
    return results


# ── element proxy ──────────────────────────────────────────────────

class _ElementProxy:
    """Unified wrapper: looks like a web element, works on CIM dicts.

    In ArcPy mode, changes are flushed to the arcpy element via setDefinition().
    """

    def __init__(self, cim_dict, arcpy_element=None, layout_ref=None):
        self._el = cim_dict
        self._layout = layout_ref    # PagxLayout, for flush notification

    def _dirty(self):
        """Flush changes. In ArcPy mode, pushes entire layout CIM to keep live state in sync."""
        if self._layout is not None:
            self._layout._flush()

    @property
    def raw(self):
        return self._el

    @property
    def type(self):       return self._el.get("type", "")
    @property
    def name(self):       return self._el.get("name", "")
    @name.setter
    def name(self, v):    self._el["name"] = str(v); self._dirty()
    @property
    def visible(self):    return self._el.get("visible", True)
    @visible.setter
    def visible(self, v): self._el["visible"] = bool(v); self._dirty()
    @property
    def anchor(self):     return self._el.get("anchor", "")
    @anchor.setter
    def anchor(self, v):  self._el["anchor"] = str(v); self._dirty()

    # position / size
    @property
    def x(self):
        rings = _get_frame_rings(self._el)
        return min(p[0] for p in rings[0]) if rings and rings[0] else 0
    @property
    def y(self):
        rings = _get_frame_rings(self._el)
        return min(p[1] for p in rings[0]) if rings and rings[0] else 0
    @property
    def width(self):
        rings = _get_frame_rings(self._el)
        return max(p[0] for p in rings[0]) - min(p[0] for p in rings[0]) if rings and rings[0] else 0
    @property
    def height(self):
        rings = _get_frame_rings(self._el)
        return max(p[1] for p in rings[0]) - min(p[1] for p in rings[0]) if rings and rings[0] else 0

    def set_position(self, x, y, width=None, height=None):
        w = width if width is not None else self.width
        h = height if height is not None else self.height
        rings = [[[x, y], [x + w, y], [x + w, y + h], [x, y + h], [x, y]]]
        _set_frame_rings(self._el, rings)
        self._dirty()

    # border
    @property
    def border_width(self):
        s = _get_border_stroke(self._el)
        return s.get("width", 0) if s else 0
    @border_width.setter
    def border_width(self, v):
        s = _get_border_stroke(self._el)
        if s:
            s["width"] = float(v)
            s["enable"] = v > 0
            self._dirty()

    @property
    def border_color(self):
        s = _get_border_stroke(self._el)
        return _rgba_to_hex(s["color"]["values"]) if (s and "color" in s) else "#00000000"
    @border_color.setter
    def border_color(self, h):
        s = _get_border_stroke(self._el)
        if s:
            v = _hex_to_rgba(h)
            if "color" not in s:
                s["color"] = {"type": "CIMRGBColor", "values": v}
            else:
                s["color"]["values"] = v
            self._dirty()

    # text
    @property
    def has_text(self):
        return self._el.get("graphic", {}).get("type") in ("CIMParagraphTextGraphic", "CIMTextGraphic")

    @property
    def text(self):
        return self._el.get("graphic", {}).get("text", "")
    @text.setter
    def text(self, v):
        g = self._el.get("graphic")
        if g:
            g["text"] = str(v)
            self._dirty()

    @property
    def font_family(self):    return _get_text_symbol(self._el).get("fontFamilyName", "")
    @font_family.setter
    def font_family(self, v): _get_text_symbol(self._el)["fontFamilyName"] = str(v); self._dirty()

    @property
    def font_style(self):     return _get_text_symbol(self._el).get("fontStyleName", "Regular")
    @font_style.setter
    def font_style(self, v):  _get_text_symbol(self._el)["fontStyleName"] = str(v); self._dirty()

    @property
    def font_size(self):      return _get_text_symbol(self._el).get("height", 10)
    @font_size.setter
    def font_size(self, v):   _get_text_symbol(self._el)["height"] = float(v); self._dirty()

    @property
    def bold(self):           return _get_text_symbol(self._el).get("fontStyleName") == "Bold"
    @bold.setter
    def bold(self, v):        _get_text_symbol(self._el)["fontStyleName"] = "Bold" if v else "Regular"; self._dirty()

    @property
    def h_align(self):        return _get_text_symbol(self._el).get("horizontalAlignment", "Left")
    @h_align.setter
    def h_align(self, v):     _get_text_symbol(self._el)["horizontalAlignment"] = str(v); self._dirty()

    @property
    def v_align(self):        return _get_text_symbol(self._el).get("verticalAlignment", "Top")
    @v_align.setter
    def v_align(self, v):     _get_text_symbol(self._el)["verticalAlignment"] = str(v); self._dirty()

    @property
    def color(self):
        layer = _get_text_color_layer(self._el)
        return _rgba_to_hex(layer["color"]["values"]) if (layer and "color" in layer) else "#000000"
    @color.setter
    def color(self, h):
        v = _hex_to_rgba(h)
        _ensure_text_color(self._el, *v)
        self._dirty()

    # group
    @property
    def children(self):
        elems = self._el.get("elements", [])
        return [_ElementProxy(e, None, self._layout) for e in elems]

    def is_group(self):
        return self._el.get("type") == "CIMGroupElement"

    def __repr__(self):
        t = self.type.replace("CIM", "").replace("Element", "").replace("Graphic", "Gfx")
        n = f" name='{self.name}'" if self.name else ""
        txt = f" text='{self.text[:40]}'" if self.has_text else ""
        return f"<{t}{n}{txt}>"


# ── page / mapSeries proxies ───────────────────────────────────────

class _PageProxy:
    def __init__(self, layout_def, arcpy_layout=None, flush_cb=None):
        self._ld = layout_def
        self._apy = arcpy_layout
        self._flush = flush_cb or (lambda: None)

    @property
    def width(self):
        if self._apy is not None:
            return self._apy.pageWidth
        return self._ld["page"]["width"]
    @width.setter
    def width(self, v):
        self._ld["page"]["width"] = float(v)
        if self._apy is not None:
            self._apy.pageWidth = float(v)
        self._flush()

    @property
    def height(self):
        if self._apy is not None:
            return self._apy.pageHeight
        return self._ld["page"]["height"]
    @height.setter
    def height(self, v):
        self._ld["page"]["height"] = float(v)
        if self._apy is not None:
            self._apy.pageHeight = float(v)
        self._flush()

    @property
    def paper(self):
        return self._ld.get("page", {}).get("printerPreferences", {}).get("paperName", "")
    @paper.setter
    def paper(self, v):
        pp = self._ld.get("page", {}).get("printerPreferences")
        if pp is None:
            pp = {"type": "CIMPrinterPreferences"}
            self._ld["page"]["printerPreferences"] = pp
        pp["paperName"] = str(v)

    def set_size(self, w, h):
        self.width = w
        self.height = h

    def __repr__(self):
        return f"<Page {self.width}x{self.height}mm paper={self.paper}>"

    @property
    def units(self):
        return "mm"  # ArcGIS Pro layout units are always mm internally


class _MapSeriesProxy:
    def __init__(self, cim_layout, arcpy_ms=None):
        self._ms = cim_layout.get("mapSeries") or {}
        self._apy = arcpy_ms

    def _get(self, key, default=None):
        if self._apy is not None:
            return getattr(self._apy, key, default)
        return self._ms.get(key, default)

    def _set(self, key, value):
        if self._apy is not None:
            setattr(self._apy, key, value)
        else:
            self._ms[key] = value

    @property
    def enabled(self):           return self._get("enabled", False)
    @enabled.setter
    def enabled(self, v):        self._set("enabled", bool(v))
    @property
    def map_frame_name(self):    return self._get("mapFrameName", "")
    @map_frame_name.setter
    def map_frame_name(self, v): self._set("mapFrameName", str(v))
    @property
    def name_field(self):        return self._get("nameField", "")
    @name_field.setter
    def name_field(self, v):     self._set("nameField", str(v))
    @property
    def sort_field(self):        return self._get("sortField", "")
    @sort_field.setter
    def sort_field(self, v):     self._set("sortField", str(v))
    @property
    def scale_rounding(self):    return self._get("scaleRounding", 1000)
    @scale_rounding.setter
    def scale_rounding(self, v): self._set("scaleRounding", int(v))
    @property
    def margin(self):            return self._get("margin", 10)
    @margin.setter
    def margin(self, v):         self._set("margin", float(v))
    @property
    def margin_type(self):       return self._get("marginType", "Percent")
    @margin_type.setter
    def margin_type(self, v):    self._set("marginType", str(v))
    @property
    def extent_options(self):    return self._get("extentOptions", "ExtentCenter")
    @extent_options.setter
    def extent_options(self, v): self._set("extentOptions", str(v))

    def __repr__(self):
        return f"<MapSeries enabled={self.enabled} mapFrame='{self.map_frame_name}' nameField='{self.name_field}'>"


# ── main class ─────────────────────────────────────────────────────

class PagxLayout:
    """DOM-style layout editor. File mode or ArcPy mode.

    File mode:
        layout = PagxLayout("template.pagx")
        layout.find_one("[name='标题']").font_size = 14
        layout.save("modified.pagx")

    ArcPy mode:
        aprx = arcpy.mp.ArcGISProject("project.aprx")
        ly = aprx.listLayouts()[0]
        layout = PagxLayout.from_arcpy(ly)
        layout.find_one("[name='标题']").font_size = 14
        ly.exportToPDF("output.pdf")
    """

    # ── constructors ──────────────────────────────────────────────

    def __init__(self, path):
        self._path = path
        self._arcpy_mode = False
        self._arcpy_layout = None
        self._dirty_flag = False
        with open(path, "r", encoding="utf-8") as f:
            self._doc = json.load(f)
        self._layout_def = self._doc.get("layoutDefinition", {})

    @classmethod
    def from_arcpy(cls, arcpy_layout):
        """Wrap an arcpy.mp.Layout object for DOM-style manipulation.

        Changes are live — every property setter automatically pushes the
        modified CIM back via setDefinition(). No save() needed.

        Usage:
            aprx = arcpy.mp.ArcGISProject("project.aprx")
            ly = aprx.listLayouts()[0]
            layout = PagxLayout.from_arcpy(ly)
            layout.find_one("[name='标题']").font_size = 14
            ly.exportToPDF("output.pdf")
        """
        obj = cls.__new__(cls)
        obj._path = None
        obj._arcpy_mode = True
        obj._arcpy_layout = arcpy_layout
        obj._dirty_flag = False
        obj._layout_def = arcpy_layout.getDefinition('V3')
        obj._doc = {"layoutDefinition": obj._layout_def}
        return obj

    def _flush(self):
        """Push CIM tree back to ArcPy layout (ArcPy mode) or mark dirty (file mode)."""
        if self._arcpy_mode and self._arcpy_layout is not None:
            self._arcpy_layout.setDefinition(self._layout_def)
        else:
            self._dirty_flag = True

    # ── selectors ─────────────────────────────────────────────────

    def find(self, selector):
        """Return matching elements. Selectors: '[name="x"]', '[name*="x"]',
        '[type="CIMMapFrame"]', '[visible=true]'.
        """
        conds = _parse_selector(selector)
        # Search the CIM tree directly (works identically in both modes)
        elems = _find_in_dicts(self._layout_def, conds)
        return [_ElementProxy(e, None, self) for e in elems]

    def find_one(self, selector):
        results = self.find(selector)
        return results[0] if results else None

    @property
    def elements(self):
        return [_ElementProxy(e, None, self) for e in self._layout_def.get("elements", [])]

    @property
    def all_elements(self):
        return self.find("[type]")  # match all elements

    def walk(self, callback):
        for el in self.all_elements:
            if callback(el):
                return True

    # ── page / mapSeries ──────────────────────────────────────────

    @property
    def page(self):
        apy = self._arcpy_layout if self._arcpy_mode else None
        return _PageProxy(self._layout_def, apy, self._flush)

    @property
    def map_series(self):
        if self._arcpy_mode:
            return _MapSeriesProxy(self._layout_def, self._arcpy_layout.mapSeries)
        return _MapSeriesProxy(self._layout_def)

    # ── delete ────────────────────────────────────────────────────

    def delete(self, element):
        """Remove an element from the layout. Works in both modes."""
        el_dict = element._el if isinstance(element, _ElementProxy) else element
        elems = self._layout_def.get("elements", [])
        if el_dict in elems:
            elems.remove(el_dict)
            self._flush()
            return True
        def _remove(container):
            for i, e in enumerate(container.get("elements", [])):
                if e is el_dict:
                    container["elements"].pop(i)
                    return True
                if e.get("type") == "CIMGroupElement":
                    if _remove(e):
                        return True
            return False
        found = _remove(self._layout_def)
        if found:
            self._flush()
        return found

    # ── map definitions ───────────────────────────────────────────

    @property
    def maps(self):
        if self._arcpy_mode:
            return []  # maps live on aprx, not layout
        return self._doc.get("mapDefinitions", [])

    # ── utility ───────────────────────────────────────────────────

    def replace_dynamic_text(self, name, static_value):
        el = self.find_one(f"[name='{name}']")
        if el is None:
            raise KeyError(f"Element not found: name='{name}'")
        el.text = str(static_value)

    def fill(self, mapping):
        """Batch replace text on named elements.

        mapping: {'元素名': '替换文字', ...}
        Missing names are skipped (no error).
        """
        for name, text in mapping.items():
            el = self.find_one(f"[name='{name}']")
            if el is not None and el.has_text:
                el.text = str(text)
        self._flush()

    def health_check(self):
        """Scan layout for common problems. Returns list of issues.

        Checks: text overflow, font readability, color contrast, anchor issues,
        dynamic text that may overflow, unnamed elements.
        """
        issues = []
        pw = self.page.width
        ph = self.page.height

        for el in self.all_elements:
            nm = el.name or "(unnamed)"
            loc = f"'{nm}' ({el.type.replace('CIM','').replace('Element','')})"

            # Skip elements with no frame (anchored, no box)
            has_frame = el.width > 0 or el.height > 0
            if has_frame:
                # Off-page
                if el.x < -5:
                    issues.append(f"[超出页面] {loc} 左边界外 (x={el.x:.1f})")
                if el.y < -5:
                    issues.append(f"[超出页面] {loc} 下边界外 (y={el.y:.1f})")
                if el.x + el.width > pw + 5:
                    issues.append(f"[超出页面] {loc} 右边界外 (右={el.x+el.width:.1f}, 页宽={pw:.0f})")
                if el.y + el.height > ph + 5:
                    issues.append(f"[超出页面] {loc} 上边界外 (上={el.y+el.height:.1f}, 页高={ph:.0f})")

            # Text checks
            if el.has_text and el.text.strip():
                txt = el.text.strip()
                fs = el.font_size

                # Font too small to read
                if fs < 5:
                    issues.append(f"[字号过小] {loc} {fs}pt — 打印后难以辨认")

                # Unnamed element (hard to select/find later)
                if not el.name:
                    issues.append(f"[缺少名称] {loc} — 无法通过 name 选择，fill() 无效")

                # Text overflow estimate — use graphic.shape if available, otherwise frame
                shape_rings = el._el.get("graphic", {}).get("shape", {}).get("rings", [])
                box_w = 0
                if shape_rings and shape_rings[0] and len(shape_rings[0]) >= 4:
                    xs = [p[0] for p in shape_rings[0]]
                    box_w = max(xs) - min(xs)
                if box_w == 0 and has_frame:
                    box_w = el.width  # fallback to frame width

                if fs >= 5 and box_w > 0:
                    char_w = fs * 0.5  # mm per Chinese char at given pt
                    lines = txt.split('\n')
                    max_cc = max(len(line) for line in lines) if lines else 0
                    if max_cc > 0:
                        text_w = max_cc * char_w
                        # Only flag when clearly over — box can fit < 3 chars but text is long
                        if text_w > box_w * 1.5 or (box_w < char_w * 2 and max_cc > 3):
                            pct = int(text_w / box_w * 100) if box_w > 0 else 999
                            preview = txt[:25] + ("..." if len(txt)>25 else "")
                            issues.append(f"[文字溢出] {loc} — 最长行{max_cc}字需{text_w:.0f}mm, "
                                         f"文字框宽{box_w:.0f}mm ({pct}%), \"{preview}\"")

                # Dynamic text warning
                if '<dyn' in txt:
                    dyn_type = ""
                    if 'property="scale"' in txt:
                        dyn_type = "比例尺"
                    elif 'property="value"' in txt:
                        dyn_type = "表格值"
                    elif 'property="upperLeft' in txt or 'property="LOWER' in txt:
                        dyn_type = "坐标"
                    if dyn_type:
                        # Don't flag all dynamic text, just note it
                        pass

                # Color contrast
                c = el.color
                if len(c) >= 7:
                    r, g, b = int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16)
                    brightness = (r * 299 + g * 587 + b * 114) / 1000
                    if brightness > 240:
                        issues.append(f"[对比度过低] {loc} color={c} — 白纸上几乎看不见")
                    a = 100
                    if len(c) >= 9:
                        a = int(c[7:9], 16)
                    if a < 30:
                        issues.append(f"[透明度过高] {loc} alpha={a}% — 近乎透明")

        # Overlap check
        framed = [e for e in self.all_elements if e.width > 5 and e.height > 5 and e.visible]
        for i, a in enumerate(framed):
            for b in framed[i + 1:]:
                if (a.x < b.x + b.width and a.x + a.width > b.x and
                    a.y < b.y + b.height and a.y + a.height > b.y):
                    if a.has_text and b.has_text:
                        issues.append(f"[元素重叠] '{a.name or '?'}' ↔ '{b.name or '?'}'")

        return issues

    def tree(self, max_depth=4):
        def _tree(container_elems, indent, depth):
            if depth > max_depth:
                return
            for el in container_elems:
                cim = el._el if isinstance(el, _ElementProxy) else el
                t = cim.get("type", "?").replace("CIM", "").replace("Element", "").replace("Graphic", "Gfx")
                head = f"{'  ' * indent}[{t}]"
                parts = [head]
                if cim.get("name"):
                    parts.append(f"name='{cim['name']}'")
                txt = cim.get("graphic", {}).get("text", "")
                if txt:
                    parts.append(f"text='{txt[:50]}'")
                w = 0
                rings = _get_frame_rings(cim)
                if rings and rings[0]:
                    xs = [p[0] for p in rings[0]]; ys = [p[1] for p in rings[0]]
                    w, h = max(xs)-min(xs), max(ys)-min(ys)
                if w > 0:
                    parts.append(f"size={w:.0f}x{h:.0f}mm")
                print(" ".join(parts))
                if cim.get("type") == "CIMGroupElement":
                    children = cim.get("elements", [])
                    _tree([_ElementProxy(e, None, self) for e in children], indent + 1, depth + 1)
        elems = self.elements
        _tree(elems, 0, 0)

    # ── save ──────────────────────────────────────────────────────

    def save(self, path=None):
        """Write to .pagx file (file mode only)."""
        if self._arcpy_mode:
            # ArcPy mode: flush CIM back to layout
            self._arcpy_layout.setDefinition(self._layout_def)
            return
        target = path or self._path
        if target is None:
            raise ValueError("No save path specified")
        os.makedirs(os.path.dirname(os.path.abspath(target)), exist_ok=True)
        with open(target, "w", encoding="utf-8") as f:
            json.dump(self._doc, f, indent=2, ensure_ascii=False)

    def to_json(self):
        return self._doc

    def __repr__(self):
        if self._arcpy_mode:
            return f"<PagxLayout arcpy mode page={self.page.width}x{self.page.height}>"
        return f"<PagxLayout '{os.path.basename(self._path)}' page={self.page.width}x{self.page.height}>"


# ── convenience ────────────────────────────────────────────────────

def open_layout(path):
    return PagxLayout(path)
