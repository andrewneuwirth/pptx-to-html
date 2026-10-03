#!/usr/bin/env python3
"""pptx_to_html.py - fast, high-fidelity PowerPoint (.pptx) -> self-contained HTML.

Every slide becomes real HTML + inline SVG (selectable text, vector shapes,
embedded images, tables, charts, SmartArt, EMF vector art) that scales to the
width of its container. All slides are rendered top to bottom.

No LibreOffice, no network access, no external files.
Dependencies: python-pptx (which itself brings lxml, Pillow, XlsxWriter). Python 3.9+.

    from pptx_to_html import convert
    html = convert("deck.pptx")                 # path, bytes, or file-like
    html = convert(data, full_document=False)   # embeddable fragment

    python pptx_to_html.py deck.pptx -o deck.html

Preset shape geometry is the ECMA-376 presetShapeDefinitions table, embedded
compressed at the bottom of this file.
"""
from __future__ import annotations

import base64
import colorsys
import copy
import html as _html
import io
import json
import math
import re
import struct
import sys
import time
import zlib
from typing import Dict, List, Optional, Tuple

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

__all__ = ["convert", "convert_file", "render_slides", "slide_text"]

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
R_EMBED = "{%s}embed" % R_NS
R_ID = "{%s}id" % R_NS
R_DM = "{%s}dm" % R_NS
EMU_PT = 12700.0
DEFAULT_W, DEFAULT_H = 12192000, 6858000


# --------------------------------------------------------------------------- utils

def _ln(el):
    t = el.tag
    if not isinstance(t, str):
        return ""
    i = t.rfind("}")
    return t[i + 1:] if i >= 0 else t


def kid(el, *names):
    if el is None:
        return None
    for c in el:
        if _ln(c) in names:
            return c
    return None


def kids(el, name):
    if el is None:
        return []
    return [c for c in el if _ln(c) == name]


def kids_all(el):
    return [c for c in el if _ln(c)] if el is not None else []


def path(el, *names):
    for n in names:
        el = kid(el, n)
        if el is None:
            return None
    return el


def _int(v, default=0):
    if v is None:
        return default
    try:
        return int(v)
    except (TypeError, ValueError):
        try:
            return int(float(str(v).rstrip("%")))
        except (TypeError, ValueError):
            return default


def iattr(el, name, default=0):
    return default if el is None else _int(el.get(name), default)


def battr(el, name, default=False):
    v = el.get(name) if el is not None else None
    if v is None:
        return default
    return v in ("1", "true", "on")


def num(v, p=3):
    s = "%.*f" % (p, v)
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    return "0" if s in ("", "-0") else s


def esc(s):
    """Escape for text and double-quoted attributes without ever emitting an entity that decodes to a
    double quote: Open WebUI 0.11 drops embeds containing &quot;/&#34; (open-webui issue #28085)."""
    return _html.escape(s, quote=False).replace('"', "'")


_CTRL = re.compile("[\x00-\x08\x0c\x0e-\x1f￾￿]")


def clean(s):
    return _CTRL.sub("", s or "")


def qa(name):
    return "{%s}%s" % (A_NS, name)


def _text_of(el):
    return "".join(t.text or "" for t in el.iter(qa("t"))) if el is not None else ""


# ------------------------------------------------------------------ geometry engine

_PRESET_CACHE = None


def _presets():
    global _PRESET_CACHE
    if _PRESET_CACHE is None:
        _PRESET_CACHE = json.loads(zlib.decompress(base64.b64decode("".join(_PRESET_B64.split()))))
    return _PRESET_CACHE


def _rad(a):
    return math.radians(a / 60000.0)


class _Guides:
    def __init__(self, w, h):
        ss, ls = min(w, h), max(w, h)
        v = {"w": w, "h": h, "l": 0.0, "t": 0.0, "r": w, "b": h, "hc": w / 2, "vc": h / 2,
             "ss": ss, "ls": ls, "cd2": 10800000.0, "cd4": 5400000.0, "cd8": 2700000.0,
             "3cd4": 16200000.0, "3cd8": 8100000.0, "5cd8": 13500000.0, "7cd8": 18900000.0}
        for n in (2, 3, 4, 5, 6, 8, 10, 12, 16, 32):
            v["wd%d" % n] = w / n
            v["hd%d" % n] = h / n
            v["ssd%d" % n] = ss / n
        self.v = v

    def val(self, tok):
        if tok is None:
            return 0.0
        try:
            return self.v[tok]
        except KeyError:
            try:
                return float(tok)
            except ValueError:
                return 0.0

    def ev(self, fmla):
        p = fmla.split()
        if not p:
            return 0.0
        op = p[0]
        a = [self.val(t) for t in p[1:]] + [0.0, 0.0, 0.0]
        x, y, z = a[0], a[1], a[2]
        try:
            if op == "val":
                return x
            if op == "*/":
                return x * y / z if z else 0.0
            if op == "+-":
                return x + y - z
            if op == "+/":
                return (x + y) / z if z else 0.0
            if op == "?:":
                return y if x > 0 else z
            if op == "abs":
                return abs(x)
            if op == "at2":
                return math.degrees(math.atan2(y, x)) * 60000.0
            if op == "cat2":
                return x * math.cos(math.atan2(z, y))
            if op == "sat2":
                return x * math.sin(math.atan2(z, y))
            if op == "cos":
                return x * math.cos(_rad(y))
            if op == "sin":
                return x * math.sin(_rad(y))
            if op == "tan":
                return x * math.tan(_rad(y))
            if op == "max":
                return max(x, y)
            if op == "min":
                return min(x, y)
            if op == "mod":
                return math.sqrt(x * x + y * y + z * z)
            if op == "pin":
                return x if y < x else (z if y > z else y)
            if op == "sqrt":
                return math.sqrt(max(x, 0.0))
        except (ValueError, OverflowError, ZeroDivisionError):
            return 0.0
        return 0.0

    def define(self, pairs, override=None):
        for name, f in pairs:
            if override and name in override:
                self.v[name] = self.ev(override[name])
            else:
                self.v[name] = self.ev(f)


def _spec_from_xml(geom):
    """custGeom element -> the same compact dict format as the embedded presets."""
    def gds(p):
        return [[g.get("name"), g.get("fmla")] for g in p.iter(qa("gd"))] if p is not None else []

    def pts(c):
        out = []
        for q in c.iter(qa("pt")):
            out += [q.get("x"), q.get("y")]
        return out

    spec = {"av": gds(kid(geom, "avLst")), "gd": gds(kid(geom, "gdLst")), "p": []}
    r = kid(geom, "rect")
    if r is not None:
        spec["rect"] = [r.get("l"), r.get("t"), r.get("r"), r.get("b")]
    for p in kids(kid(geom, "pathLst"), "path"):
        cmds = []
        for c in p:
            n = _ln(c)
            if n == "moveTo":
                cmds.append(["M"] + pts(c))
            elif n == "lnTo":
                cmds.append(["L"] + pts(c))
            elif n == "cubicBezTo":
                cmds.append(["C"] + pts(c))
            elif n == "quadBezTo":
                cmds.append(["Q"] + pts(c))
            elif n == "arcTo":
                cmds.append(["A", c.get("wR"), c.get("hR"), c.get("stAng"), c.get("swAng")])
            elif n == "close":
                cmds.append(["Z"])
        pd = {"c": cmds}
        for k in ("w", "h", "fill", "stroke"):
            if p.get(k) is not None:
                pd[k] = p.get(k)
        spec["p"].append(pd)
    return spec


class Geometry:
    __slots__ = ("paths", "rect", "polys")

    def __init__(self, paths, rect, polys):
        self.paths = paths  # [(svg_d, fill_mode, stroke_bool)]
        self.rect = rect    # text rect (l, t, r, b) in EMU inside the shape
        self.polys = polys  # flattened point lists per path (for clip-path)


_RECT_SPEC = {"av": [], "gd": [], "p": [{"c": [["M", "l", "t"], ["L", "r", "t"], ["L", "r", "b"], ["L", "l", "b"], ["Z"]]}]}


def build_geometry(spPr_list, w, h):
    """spPr_list: candidate spPr elements, highest priority first."""
    spec, override = None, None
    for spPr in spPr_list:
        if spPr is None:
            continue
        cust = kid(spPr, "custGeom")
        if cust is not None:
            spec = _spec_from_xml(cust)
            break
        prst = kid(spPr, "prstGeom")
        if prst is not None:
            spec = _presets().get(prst.get("prst") or "rect")
            override = {g.get("name"): g.get("fmla") for g in prst.iter(qa("gd"))}
            break
    if spec is None:
        spec = _RECT_SPEC
    w = max(float(w), 1.0)
    h = max(float(h), 1.0)
    g = _Guides(w, h)
    g.define(spec.get("av", []), override)
    g.define(spec.get("gd", []))
    paths, polys = [], []
    for p in spec.get("p", []):
        pw, ph = float(p.get("w") or 0), float(p.get("h") or 0)
        sx = w / pw if pw else 1.0
        sy = h / ph if ph else 1.0
        d, poly = [], []
        cx = cy = 0.0
        start = (0.0, 0.0)
        for c in p["c"]:
            k = c[0]
            if k == "M":
                cx, cy = g.val(c[1]) * sx, g.val(c[2]) * sy
                start = (cx, cy)
                d.append("M%s %s" % (num(cx, 1), num(cy, 1)))
                poly.append((cx, cy))
            elif k == "L":
                cx, cy = g.val(c[1]) * sx, g.val(c[2]) * sy
                d.append("L%s %s" % (num(cx, 1), num(cy, 1)))
                poly.append((cx, cy))
            elif k == "C" and len(c) >= 7:
                q = [g.val(c[i]) * (sx if i % 2 else sy) for i in range(1, 7)]
                d.append("C" + " ".join(num(v, 1) for v in q))
                for s in range(1, 9):
                    t = s / 8.0
                    mt = 1 - t
                    poly.append((mt ** 3 * cx + 3 * mt * mt * t * q[0] + 3 * mt * t * t * q[2] + t ** 3 * q[4],
                                 mt ** 3 * cy + 3 * mt * mt * t * q[1] + 3 * mt * t * t * q[3] + t ** 3 * q[5]))
                cx, cy = q[4], q[5]
            elif k == "Q" and len(c) >= 5:
                q = [g.val(c[i]) * (sx if i % 2 else sy) for i in range(1, 5)]
                d.append("Q" + " ".join(num(v, 1) for v in q))
                poly.append((q[2], q[3]))
                cx, cy = q[2], q[3]
            elif k == "A":
                cx, cy = _arc(d, poly, cx, cy, g.val(c[1]) * sx, g.val(c[2]) * sy, g.val(c[3]), g.val(c[4]))
            elif k == "Z":
                d.append("Z")
                cx, cy = start
        paths.append(("".join(d), p.get("fill", "norm"), p.get("stroke", "true") not in ("false", "0")))
        polys.append(poly)
    rect = (0.0, 0.0, w, h)
    if spec.get("rect"):
        r = spec["rect"]
        rect = (g.val(r[0]), g.val(r[1]), g.val(r[2]), g.val(r[3]))
    return Geometry(paths, rect, polys)


def _arc(d, poly, x0, y0, wR, hR, st, sw):
    if wR <= 0 or hR <= 0:
        return x0, y0

    def par(a):  # visual angle -> parametric angle on the ellipse
        return math.atan2(wR * math.sin(a), hR * math.cos(a))

    a1 = _rad(st)
    t1 = par(a1)
    ccx, ccy = x0 - wR * math.cos(t1), y0 - hR * math.sin(t1)
    total = _rad(sw)
    steps = max(1, int(math.ceil(abs(total) / (math.pi / 2) - 1e-9)))
    x, y = x0, y0
    for i in range(1, steps + 1):
        a0 = a1 + total * (i - 1) / steps
        a = a1 + total * i / steps
        for s in (0.2, 0.4, 0.6, 0.8, 1.0):
            tt = par(a0 + (a - a0) * s)
            poly.append((ccx + wR * math.cos(tt), ccy + hR * math.sin(tt)))
        t = par(a)
        x, y = ccx + wR * math.cos(t), ccy + hR * math.sin(t)
        d.append("A%s %s 0 0 %d %s %s" % (num(wR, 1), num(hR, 1), 1 if sw > 0 else 0, num(x, 1), num(y, 1)))
    return x, y


# ----------------------------------------------------------------------------- color

class Clr:
    __slots__ = ("r", "g", "b", "a")

    def __init__(self, r, g, b, a=1.0):
        self.r, self.g, self.b, self.a = r, g, b, a

    def css(self, alpha_mult=1.0):
        r, g, b = (int(round(max(0.0, min(1.0, v)) * 255)) for v in (self.r, self.g, self.b))
        a = max(0.0, min(1.0, self.a * alpha_mult))
        if a >= 0.999:
            return "#%02x%02x%02x" % (r, g, b)
        return "rgba(%d,%d,%d,%s)" % (r, g, b, num(a, 3))

    def solid(self):
        return Clr(self.r, self.g, self.b, 1.0)

    def lum(self):
        return 0.2126 * self.r + 0.7152 * self.g + 0.0722 * self.b

    def mix(self, other, t):
        return Clr(self.r + (other.r - self.r) * t, self.g + (other.g - self.g) * t,
                   self.b + (other.b - self.b) * t, self.a)


def _hexclr(s, default="000000"):
    s = (s or default).strip().lstrip("#")
    if len(s) != 6:
        s = default
    try:
        return Clr(int(s[0:2], 16) / 255.0, int(s[2:4], 16) / 255.0, int(s[4:6], 16) / 255.0)
    except ValueError:
        return Clr(0, 0, 0)


def _to_lin(c):
    return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4


def _to_srgb(c):
    c = max(0.0, min(1.0, c))
    return c * 12.92 if c <= 0.0031308 else 1.055 * c ** (1 / 2.4) - 0.055


_PRST_COLORS = {
    "black": "000000", "white": "FFFFFF", "red": "FF0000", "green": "008000", "blue": "0000FF",
    "yellow": "FFFF00", "cyan": "00FFFF", "magenta": "FF00FF", "gray": "808080", "grey": "808080",
    "darkGray": "A9A9A9", "lightGray": "D3D3D3", "orange": "FFA500", "purple": "800080",
    "navy": "000080", "maroon": "800000", "olive": "808000", "teal": "008080", "silver": "C0C0C0",
    "lime": "00FF00", "brown": "A52A2A", "pink": "FFC0CB", "gold": "FFD700", "darkBlue": "00008B",
    "darkRed": "8B0000", "darkGreen": "006400", "ltGray": "D3D3D3", "dkGray": "A9A9A9",
}
_SYS_COLORS = {"windowText": "000000", "window": "FFFFFF", "btnFace": "F0F0F0", "btnText": "000000",
               "highlight": "0078D7", "highlightText": "FFFFFF", "grayText": "6D6D6D"}
COLOR_TAGS = ("srgbClr", "schemeClr", "sysClr", "prstClr", "scrgbClr", "hslClr")
_DEFAULT_CLRMAP = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2", "accent1": "accent1",
                   "accent2": "accent2", "accent3": "accent3", "accent4": "accent4", "accent5": "accent5",
                   "accent6": "accent6", "hlink": "hlink", "folHlink": "folHlink"}


def color_el(parent):
    return kid(parent, *COLOR_TAGS)


def apply_mods(c, cel):
    r, g, b, a = c.r, c.g, c.b, c.a
    for m in cel:
        n = _ln(m)
        if not n:
            continue
        v = _int(m.get("val"), 0) / 100000.0
        if n == "alpha":
            a = v
        elif n == "alphaMod":
            a *= v
        elif n == "alphaOff":
            a += v
        elif n in ("lumMod", "lumOff", "satMod", "satOff", "hueMod", "hueOff", "hue", "sat", "lum"):
            hh, ll, ss = colorsys.rgb_to_hls(r, g, b)
            if n == "lumMod":
                ll *= v
            elif n == "lumOff":
                ll += v
            elif n == "lum":
                ll = v
            elif n == "satMod":
                ss *= v
            elif n == "satOff":
                ss += v
            elif n == "sat":
                ss = v
            elif n == "hueMod":
                hh = (hh * v) % 1.0
            elif n == "hueOff":
                hh = (hh + _int(m.get("val")) / 21600000.0) % 1.0
            elif n == "hue":
                hh = (_int(m.get("val")) / 21600000.0) % 1.0
            r, g, b = colorsys.hls_to_rgb(hh, max(0.0, min(1.0, ll)), max(0.0, min(1.0, ss)))
        elif n == "tint":
            r, g, b = (_to_srgb(_to_lin(x) * v + (1 - v)) for x in (r, g, b))
        elif n == "shade":
            r, g, b = (_to_srgb(_to_lin(x) * v) for x in (r, g, b))
        elif n in ("comp",):
            hh, ll, ss = colorsys.rgb_to_hls(r, g, b)
            r, g, b = colorsys.hls_to_rgb((hh + 0.5) % 1.0, ll, ss)
        elif n == "inv":
            r, g, b = 1 - r, 1 - g, 1 - b
        elif n == "gray":
            y = 0.299 * r + 0.587 * g + 0.114 * b
            r = g = b = y
        elif n == "redMod":
            r *= v
        elif n == "greenMod":
            g *= v
        elif n == "blueMod":
            b *= v
        elif n == "redOff":
            r += v
        elif n == "greenOff":
            g += v
        elif n == "blueOff":
            b += v
    clamp = lambda x: max(0.0, min(1.0, x))  # noqa: E731
    return Clr(clamp(r), clamp(g), clamp(b), clamp(a))


# ----------------------------------------------------------------------------- theme

_DEFAULT_THEME_COLORS = {
    "dk1": "000000", "lt1": "FFFFFF", "dk2": "44546A", "lt2": "E7E6E6", "accent1": "4472C4",
    "accent2": "ED7D31", "accent3": "A5A5A5", "accent4": "FFC000", "accent5": "5B9BD5",
    "accent6": "70AD47", "hlink": "0563C1", "folHlink": "954F72"}


class Theme:
    def __init__(self, root=None, part=None):
        self.part = part
        self.colors = {k: _hexclr(v) for k, v in _DEFAULT_THEME_COLORS.items()}
        self.major = {"latin": "Calibri Light"}
        self.minor = {"latin": "Calibri"}
        self.fills, self.lines, self.effects, self.bgfills = [], [], [], []
        if root is None:
            return
        el = kid(root, "themeElements")
        cs = kid(el, "clrScheme")
        for c in (cs if cs is not None else []):
            ce = color_el(c)
            if ce is None:
                continue
            if _ln(ce) == "srgbClr":
                self.colors[_ln(c)] = _hexclr(ce.get("val"))
            elif _ln(ce) == "sysClr":
                self.colors[_ln(c)] = _hexclr(ce.get("lastClr") or _SYS_COLORS.get(ce.get("val"), "000000"))
        fs = kid(el, "fontScheme")
        for which, store in (("majorFont", self.major), ("minorFont", self.minor)):
            f = kid(fs, which)
            for script in ("latin", "ea", "cs"):
                e = kid(f, script)
                if e is not None and e.get("typeface"):
                    store[script] = e.get("typeface")
        fmt = kid(el, "fmtScheme")
        self.fills = [c for c in kids_all(kid(fmt, "fillStyleLst"))]
        self.lines = [c for c in kids_all(kid(fmt, "lnStyleLst"))]
        self.effects = [c for c in kids_all(kid(fmt, "effectStyleLst"))]
        self.bgfills = [c for c in kids_all(kid(fmt, "bgFillStyleLst"))]

    def font(self, typeface):
        if not typeface or not typeface.startswith("+"):
            return typeface
        src = self.major if typeface.startswith("+mj") else self.minor
        script = typeface[-2:]
        script = {"lt": "latin", "ea": "ea", "cs": "cs"}.get(script, "latin")
        return src.get(script) or src.get("latin")


# ------------------------------------------------------------------------------ fonts

_FONT_STACKS = {
    "calibri": "Calibri,Carlito,'Segoe UI','Helvetica Neue',Arial,sans-serif",
    "calibri light": "'Calibri Light',Calibri,Carlito,'Segoe UI Light','Helvetica Neue',Arial,sans-serif",
    "aptos": "Aptos,'Segoe UI','Helvetica Neue',Arial,sans-serif",
    "aptos display": "'Aptos Display',Aptos,'Segoe UI','Helvetica Neue',Arial,sans-serif",
    "aptos narrow": "'Aptos Narrow','Arial Narrow','Helvetica Neue',Arial,sans-serif",
    "arial": "Arial,'Liberation Sans',Helvetica,sans-serif",
    "arial narrow": "'Arial Narrow','Liberation Sans Narrow',Arial,sans-serif",
    "helvetica": "Helvetica,Arial,sans-serif",
    "helvetica neue": "'Helvetica Neue',Helvetica,Arial,sans-serif",
    "times new roman": "'Times New Roman',Tinos,Times,serif",
    "times": "Times,'Times New Roman',serif",
    "cambria": "Cambria,Caladea,Georgia,serif",
    "cambria math": "'Cambria Math',Cambria,'STIX Two Math','Times New Roman',serif",
    "georgia": "Georgia,'Times New Roman',serif",
    "garamond": "Garamond,'EB Garamond','Adobe Garamond Pro',Georgia,serif",
    "book antiqua": "'Book Antiqua',Palatino,'Palatino Linotype',Georgia,serif",
    "palatino linotype": "'Palatino Linotype',Palatino,'Book Antiqua',Georgia,serif",
    "century gothic": "'Century Gothic','Avenir Next',Avenir,Futura,'Helvetica Neue',sans-serif",
    "gill sans mt": "'Gill Sans MT','Gill Sans','Gill Sans Nova','Helvetica Neue',sans-serif",
    "segoe ui": "'Segoe UI',system-ui,-apple-system,'Helvetica Neue',Arial,sans-serif",
    "segoe ui light": "'Segoe UI Light','Segoe UI',system-ui,'Helvetica Neue',sans-serif",
    "segoe ui semibold": "'Segoe UI Semibold','Segoe UI',system-ui,'Helvetica Neue',sans-serif",
    "tahoma": "Tahoma,Verdana,'Segoe UI',sans-serif",
    "verdana": "Verdana,Tahoma,sans-serif",
    "trebuchet ms": "'Trebuchet MS','Lucida Grande',Tahoma,sans-serif",
    "franklin gothic book": "'Franklin Gothic Book','Libre Franklin','Helvetica Neue',Arial,sans-serif",
    "franklin gothic medium": "'Franklin Gothic Medium','Libre Franklin','Helvetica Neue',Arial,sans-serif",
    "corbel": "Corbel,'Segoe UI','Helvetica Neue',sans-serif",
    "candara": "Candara,Calibri,'Segoe UI',sans-serif",
    "constantia": "Constantia,Georgia,serif",
    "tw cen mt": "'Tw Cen MT','Century Gothic',Futura,sans-serif",
    "rockwell": "Rockwell,'Roboto Slab','Courier New',serif",
    "impact": "Impact,'Arial Black',sans-serif",
    "arial black": "'Arial Black',Arial,sans-serif",
    "consolas": "Consolas,Menlo,Monaco,'Courier New',monospace",
    "courier new": "'Courier New',Courier,monospace",
    "lucida console": "'Lucida Console',Menlo,Monaco,monospace",
    "open sans": "'Open Sans','Segoe UI','Helvetica Neue',Arial,sans-serif",
    "roboto": "Roboto,'Segoe UI','Helvetica Neue',Arial,sans-serif",
    "montserrat": "Montserrat,'Segoe UI','Helvetica Neue',Arial,sans-serif",
    "lato": "Lato,'Segoe UI','Helvetica Neue',Arial,sans-serif",
    "poppins": "Poppins,'Segoe UI','Helvetica Neue',Arial,sans-serif",
    "inter": "Inter,'Segoe UI','Helvetica Neue',Arial,sans-serif",
}
_FONT_SAFE = re.compile(r"[^\w \-\.&]", re.UNICODE)


def font_stack(name, ea=None):
    name = _FONT_SAFE.sub("", name or "").strip()
    ea = _FONT_SAFE.sub("", ea or "").strip()
    if not name:
        name = "Calibri"
    key = name.lower()
    stack = _FONT_STACKS.get(key)
    if stack is None:
        if re.search(r"mono|courier|consol|code", key):
            generic = "monospace"
        elif re.search(r"serif|times|roman|garamond|georgia|cambria|book|palatino|baskerville|didot|bodoni|caslon|minion", key) \
                and "sans" not in key:
            generic = "serif"
        else:
            generic = "'Helvetica Neue',Arial,sans-serif"
        stack = "'%s',%s" % (name, generic)
    if ea and ea.lower() != key:
        stack = stack.replace(",", ",'%s'," % ea, 1) if "," in stack else stack
    return stack


# Symbol/Wingdings code points -> Unicode
_SYMBOL_MAP = dict(zip(
    "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ",
    "αβχδεφγηιϕκλμνοπθρστυϖωξψζΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ"))
_SYMBOL_MAP.update({
    0x22: "∀", 0x24: "∃", 0x27: "∋", 0x2A: "∗", 0x2D: "−", 0x40: "≅", 0x5C: "∴", 0x5E: "⊥",
    0xA2: "′", 0xA3: "≤", 0xA4: "⁄", 0xA5: "∞", 0xA6: "ƒ", 0xA7: "♣", 0xA8: "♦", 0xA9: "♥", 0xAA: "♠",
    0xAB: "↔", 0xAC: "←", 0xAD: "↑", 0xAE: "→", 0xAF: "↓", 0xB0: "°", 0xB1: "±", 0xB2: "″", 0xB3: "≥",
    0xB4: "×", 0xB5: "∝", 0xB6: "∂", 0xB7: "•", 0xB8: "÷", 0xB9: "≠", 0xBA: "≡", 0xBB: "≈", 0xBC: "…",
    0xBF: "↵", 0xC0: "ℵ", 0xC1: "ℑ", 0xC2: "ℜ", 0xC3: "℘", 0xC4: "⊗", 0xC5: "⊕", 0xC6: "∅", 0xC7: "∩",
    0xC8: "∪", 0xC9: "⊃", 0xCA: "⊇", 0xCB: "⊄", 0xCC: "⊂", 0xCD: "⊆", 0xCE: "∈", 0xCF: "∉", 0xD0: "∠",
    0xD1: "∇", 0xD2: "®", 0xD3: "©", 0xD4: "™", 0xD5: "∏", 0xD6: "√", 0xD7: "⋅", 0xD8: "¬", 0xD9: "∧",
    0xDA: "∨", 0xDB: "⇔", 0xDC: "⇐", 0xDD: "⇑", 0xDE: "⇒", 0xDF: "⇓", 0xE0: "◊", 0xE1: "〈", 0xE5: "∑",
    0xE6: "⎛", 0xE7: "⎜", 0xE8: "⎝", 0xE9: "⎡", 0xEA: "⎢", 0xEB: "⎣", 0xEC: "⎧", 0xED: "⎨", 0xEE: "⎩",
    0xEF: "⎪", 0xF1: "〉", 0xF2: "∫", 0xF3: "⌠", 0xF4: "⎮", 0xF5: "⌡", 0xF6: "⎞", 0xF7: "⎟", 0xF8: "⎠",
    0xF9: "⎤", 0xFA: "⎥", 0xFB: "⎦", 0xFC: "⎫", 0xFD: "⎬", 0xFE: "⎭"})
_WINGDINGS_MAP = {
    0x6C: "●", 0x6D: "❍", 0x6E: "■", 0x6F: "□", 0x70: "◻", 0x71: "❑", 0x72: "❒", 0x73: "⬧", 0x74: "⧫",
    0x75: "◆", 0x76: "❖", 0x77: "⬥", 0x9F: "•", 0xA7: "▪", 0xA8: "◻", 0xD8: "➢", 0xE0: "→", 0xE8: "➔",
    0xF0: "⇨", 0xFB: "✗", 0xFC: "✓", 0xFD: "☒", 0xFE: "☑", 0x4A: "☺", 0x4C: "☹", 0x46: "☞", 0x41: "✌",
    0xA1: "○", 0xA2: "◯", 0xA4: "◉", 0xAB: "★", 0xB2: "✧", 0x3F: "✍", 0x2A: "✉", 0x28: "☎",
    0xDF: "←", 0xE1: "↑", 0xE2: "↓", 0xD5: "⌫", 0xEF: "⇦", 0xF1: "⇧", 0xF2: "⇩"}


def map_symbol_text(text, face):
    f = (face or "").lower()
    if f.startswith("wingdings"):
        table = _WINGDINGS_MAP
    elif f in ("symbol",):
        table = _SYMBOL_MAP
    else:
        return text
    out = []
    for ch in text:
        o = ord(ch)
        if 0xF000 <= o <= 0xF0FF:
            o -= 0xF000
        out.append(table.get(o) or table.get(chr(o)) or (chr(o) if 0x20 <= o < 0x7F and table is _SYMBOL_MAP else "•" if table is _WINGDINGS_MAP else chr(o)))
    return "".join(out)


# ------------------------------------------------------------------- fills & lines

_FILL_TAGS = ("noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill")
NONE_FILL = ("none",)

_DASHES = {
    "solid": None, "dot": (1, 3), "sysDot": (1, 1), "dash": (4, 3), "sysDash": (3, 1),
    "lgDash": (8, 3), "dashDot": (4, 3, 1, 3), "sysDashDot": (3, 1, 1, 1),
    "lgDashDot": (8, 3, 1, 3), "lgDashDotDot": (8, 3, 1, 3, 1, 3), "sysDashDotDot": (3, 1, 1, 1, 1, 1)}
_ARROW_SIZE = {"sm": 2.0, "med": 3.0, "lg": 5.0}


class _Paint:
    """Mixin: colour, fill, line and effect resolution against a theme + colour map."""

    theme: Theme
    clrmap: Dict[str, str]

    def color(self, cel, ph=None):
        if cel is None:
            return None
        n = _ln(cel)
        if n == "srgbClr":
            base = _hexclr(cel.get("val"))
        elif n == "schemeClr":
            v = cel.get("val")
            if v == "phClr":
                base = Clr(ph.r, ph.g, ph.b, ph.a) if ph is not None else Clr(0, 0, 0)
            else:
                v = self.clrmap.get(v, v)
                c = self.theme.colors.get(v) or self.theme.colors.get(_DEFAULT_CLRMAP.get(v, v))
                base = Clr(c.r, c.g, c.b) if c else Clr(0, 0, 0)
        elif n == "sysClr":
            base = _hexclr(cel.get("lastClr") or _SYS_COLORS.get(cel.get("val"), "000000"))
        elif n == "prstClr":
            base = _hexclr(_PRST_COLORS.get(cel.get("val"), "000000"))
        elif n == "scrgbClr":
            base = Clr(*(_to_srgb(_int(cel.get(k)) / 100000.0) for k in ("r", "g", "b")))
        elif n == "hslClr":
            r, g, b = colorsys.hls_to_rgb((_int(cel.get("hue")) / 21600000.0) % 1.0,
                                          _int(cel.get("lum")) / 100000.0, _int(cel.get("sat")) / 100000.0)
            base = Clr(r, g, b)
        else:
            return None
        return apply_mods(base, cel)

    def scheme(self, name):
        c = self.theme.colors.get(self.clrmap.get(name, name))
        return Clr(c.r, c.g, c.b) if c else Clr(0, 0, 0)

    # fills ------------------------------------------------------------------
    def fill(self, parent, part, ph=None, grp=None):
        """First fill child of `parent` -> fill tuple, or None when unspecified."""
        if parent is None:
            return None
        for c in parent:
            n = _ln(c)
            if n not in _FILL_TAGS:
                continue
            if n == "noFill":
                return NONE_FILL
            if n == "solidFill":
                cl = self.color(color_el(c), ph)
                return ("solid", cl) if cl else NONE_FILL
            if n == "gradFill":
                stops = []
                for gs in kids(kid(c, "gsLst"), "gs"):
                    cl = self.color(color_el(gs), ph)
                    if cl:
                        stops.append((_int(gs.get("pos")) / 100000.0, cl))
                if not stops:
                    return None
                stops.sort(key=lambda s: s[0])
                lin, pth = kid(c, "lin"), kid(c, "path")
                if pth is not None:
                    ftr = kid(pth, "fillToRect")
                    rect = tuple(iattr(ftr, k) / 100000.0 for k in ("l", "t", "r", "b"))
                    return ("grad", stops, pth.get("path") or "circle", 0.0, rect)
                return ("grad", stops, "lin", iattr(lin, "ang") / 60000.0, None)
            if n == "blipFill":
                blip = kid(c, "blip")
                duo = None
                dt = kid(blip, "duotone")
                if dt is not None:
                    cols = [self.color(e, ph) for e in dt if _ln(e) in COLOR_TAGS]
                    if len(cols) == 2 and all(cols):
                        duo = cols
                img = self.deck.images.get(part, _blip_rid(blip), duo) if blip is not None else None
                if img is None:
                    return NONE_FILL
                alpha = iattr(kid(blip, "alphaModFix"), "amt", 100000) / 100000.0
                return ("img", img, kid(c, "tile"), kid(c, "srcRect"), alpha)
            if n == "pattFill":
                fg = self.color(color_el(kid(c, "fgClr")), ph) or Clr(0, 0, 0)
                bg = self.color(color_el(kid(c, "bgClr")), ph) or Clr(1, 1, 1)
                return ("patt", fg, bg, c.get("prst") or "pct50")
            if n == "grpFill":
                return grp or NONE_FILL
        return None

    def style_fill(self, style, part):
        ref = kid(style, "fillRef")
        if ref is None:
            return None
        idx = iattr(ref, "idx")
        phc = self.color(color_el(ref))
        lst = self.theme.bgfills if idx >= 1001 else self.theme.fills
        i = idx - 1001 if idx >= 1001 else idx - 1
        if idx == 0 or i < 0 or i >= len(lst):
            return NONE_FILL
        holder = etree.Element("x")
        holder.append(copy.deepcopy(lst[i]))
        return self.fill(holder, self.theme.part or part, phc)

    @staticmethod
    def fill_color(f):
        """Representative single colour of a fill (for text contrast / strokes)."""
        if not f or f[0] == "none":
            return None
        if f[0] == "solid":
            return f[1]
        if f[0] == "grad":
            return f[1][0][1]
        if f[0] == "patt":
            return f[1].mix(f[2], 0.5)
        return None

    def fill_css(self, f, slide_bg=False):
        """Fill tuple -> (css declarations, extra class)."""
        if not f or f[0] == "none":
            return "", ""
        if f[0] == "solid":
            return "background:%s" % f[1].css(), ""
        if f[0] == "patt":
            fg, bg = f[1], f[2]
            return "background-color:%s;background-image:radial-gradient(%s 22%%,transparent 26%%);background-size:4px 4px" % (
                bg.css(), fg.css()), ""
        if f[0] == "grad":
            stops = ",".join("%s %s%%" % (c.css(), num(p * 100, 1)) for p, c in f[1])
            if len(f[1]) == 1:
                return "background:%s" % f[1][0][1].css(), ""
            if f[2] == "lin":
                return "background:linear-gradient(%sdeg,%s)" % (num(f[3] + 90, 2), stops), ""
            l, t, r, b = f[4] or (0.5, 0.5, 0.5, 0.5)
            cx, cy = (l + 1 - r) / 2 * 100, (t + 1 - b) / 2 * 100
            shape = "circle" if f[2] == "circle" else "ellipse"
            return "background:radial-gradient(%s farthest-corner at %s%% %s%%,%s)" % (shape, num(cx, 1), num(cy, 1), stops), ""
        if f[0] == "img":
            img, tile, src, alpha = f[1], f[2], f[3], f[4]
            if tile is not None:
                sx, sy = iattr(tile, "sx", 100000) / 100000.0, iattr(tile, "sy", 100000) / 100000.0
                pw, ph = img.size()
                css = "background-repeat:repeat;background-size:%s %s" % (
                    self.u(pw * 9525 * sx), self.u(ph * 9525 * sy))
            else:
                css = "background-repeat:no-repeat;background-size:100% 100%"
                if src is not None:
                    css = _crop_css(src)
            if alpha < 0.999 and not slide_bg:
                css += ";opacity:%s" % num(alpha, 3)
            return css, img.cls
        return "", ""

    # svg paint ----------------------------------------------------------------
    def svg_paint(self, f, defs, mode="norm"):
        """Fill tuple -> value for an SVG fill attribute (adds gradient defs)."""
        if not f or f[0] == "none" or mode == "none":
            return "none"
        if f[0] in ("solid", "patt"):
            c = self.fill_color(f)
            return _shade(c, mode).css()
        if f[0] == "grad":
            gid = self.deck.new_id("g")
            stops = "".join('<stop offset="%s" stop-color="%s"%s/>' % (
                num(p, 4), _shade(c, mode).css(1.0) if c.a >= 0.999 else _shade(c, mode).solid().css(),
                "" if c.a >= 0.999 else ' stop-opacity="%s"' % num(c.a, 3)) for p, c in f[1])
            if f[2] == "lin":
                a = math.radians(f[3])
                ca, sa = math.cos(a), math.sin(a)
                R = (abs(ca) + abs(sa)) / 2
                defs.append('<linearGradient id="%s" x1="%s" y1="%s" x2="%s" y2="%s">%s</linearGradient>' % (
                    gid, num(.5 - ca * R, 4), num(.5 - sa * R, 4), num(.5 + ca * R, 4), num(.5 + sa * R, 4), stops))
            else:
                l, t, r, b = f[4] or (0.5, 0.5, 0.5, 0.5)
                cx, cy = (l + 1 - r) / 2, (t + 1 - b) / 2
                rr = max(math.hypot(max(cx, 1 - cx), max(cy, 1 - cy)), 0.5)
                defs.append('<radialGradient id="%s" cx="%s" cy="%s" r="%s">%s</radialGradient>' % (
                    gid, num(cx, 4), num(cy, 4), num(rr, 4), stops))
            return "url(#%s)" % gid
        if f[0] == "img":
            pid = self.deck.new_id("p")
            defs.append('<pattern id="%s" patternContentUnits="objectBoundingBox" width="1" height="1">'
                        '<image href="%s" width="1" height="1" preserveAspectRatio="none"/></pattern>' % (pid, f[1].uri))
            return "url(#%s)" % pid
        return "none"

    # lines --------------------------------------------------------------------
    def line_props(self, ln, part, ph=None):
        d = {}
        if ln is None:
            return d
        if ln.get("w") is not None:
            d["w"] = _int(ln.get("w"))
        if ln.get("cap"):
            d["cap"] = ln.get("cap")
        f = self.fill(ln, part, ph)
        if f is not None:
            d["fill"] = f
        pd = kid(ln, "prstDash")
        if pd is not None:
            d["dash"] = _DASHES.get(pd.get("val"))
        cd = kid(ln, "custDash")
        if cd is not None:
            seq = []
            for ds in kids(cd, "ds"):
                seq += [iattr(ds, "d") / 100000.0, iattr(ds, "sp") / 100000.0]
            d["dash"] = tuple(seq) or None
        for j in ("round", "bevel", "miter"):
            if kid(ln, j) is not None:
                d["join"] = j
        for end in ("headEnd", "tailEnd"):
            e = kid(ln, end)
            if e is not None:
                d[end] = (e.get("type") or "none", e.get("w") or "med", e.get("len") or "med")
        return d

    def style_line(self, style, part):
        ref = kid(style, "lnRef")
        if ref is None:
            return {}
        idx = iattr(ref, "idx")
        if idx <= 0 or idx > len(self.theme.lines):
            return {}
        return self.line_props(self.theme.lines[idx - 1], part, self.color(color_el(ref)))

    def stroke_attrs(self, line, defs):
        """Merged line dict -> (svg attribute string, stroke width EMU)."""
        f = line.get("fill")
        if not f or f[0] == "none":
            return 'stroke="none"', 0
        w = max(line.get("w", 9525), 1)
        paint = self.svg_paint(f, defs) if f[0] == "grad" else (self.fill_color(f) or Clr(0, 0, 0)).css()
        a = ['stroke="%s"' % paint, 'stroke-width="%d"' % w]
        cap = {"rnd": "round", "sq": "square", "flat": "butt"}.get(line.get("cap"), "butt")
        a.append('stroke-linecap="%s"' % cap)
        a.append('stroke-linejoin="%s"' % {"round": "round", "bevel": "bevel", "miter": "miter"}.get(line.get("join"), "round"))
        dash = line.get("dash")
        if dash:
            # round/square caps extend each dash; compensate like PowerPoint does
            adj = 1.0 if cap != "butt" else 0.0
            seq = []
            for i, v in enumerate(dash):
                seq.append(max(v * w - adj * w, 0.01 * w) if i % 2 == 0 else v * w + adj * w)
            a.append('stroke-dasharray="%s"' % " ".join("%d" % v for v in seq))
        col = self.fill_color(f) or Clr(0, 0, 0)
        for end, attr in (("headEnd", "marker-start"), ("tailEnd", "marker-end")):
            spec = line.get(end)
            if spec and spec[0] != "none":
                a.append('%s="url(#%s)"' % (attr, self._marker(spec, col, defs)))
        return " ".join(a), w

    def _marker(self, spec, col, defs):
        kind, wk, lk = spec
        mw, ml = _ARROW_SIZE.get(wk, 3.0), _ARROW_SIZE.get(lk, 3.0)
        mid = self.deck.new_id("m")
        c = col.css()
        if kind == "arrow":
            body = '<path d="M1 1L9 5L1 9" fill="none" stroke="%s" stroke-width="%s" stroke-linecap="round" stroke-linejoin="round"/>' % (c, num(10.0 / mw, 2))
            ref = 8
        elif kind == "stealth":
            body = '<path d="M0 0L10 5L0 10L3 5Z" fill="%s"/>' % c
            ref = 9
        elif kind == "diamond":
            body = '<path d="M0 5L5 0L10 5L5 10Z" fill="%s"/>' % c
            ref = 5
        elif kind == "oval":
            body = '<ellipse cx="5" cy="5" rx="5" ry="5" fill="%s"/>' % c
            ref = 5
        else:  # triangle
            body = '<path d="M0 0L10 5L0 10Z" fill="%s"/>' % c
            ref = 9
        defs.append('<marker id="%s" viewBox="0 0 10 10" refX="%d" refY="5" markerWidth="%s" markerHeight="%s" '
                    'orient="auto-start-reverse" markerUnits="strokeWidth" overflow="visible">%s</marker>' % (
                        mid, ref, num(ml, 2), num(mw, 2), body))
        return mid

    # effects ------------------------------------------------------------------
    def effect_list(self, spPr_list, style):
        for spPr in spPr_list:
            el = kid(spPr, "effectLst")
            if el is not None:
                return el, None
        ref = kid(style, "effectRef")
        if ref is not None:
            idx = iattr(ref, "idx")
            if 0 < idx <= len(self.theme.effects):
                return kid(self.theme.effects[idx - 1], "effectLst"), self.color(color_el(ref))
        return None, None

    def shadow_css(self, effects, ph=None):
        sh = kid(effects, "outerShdw")
        if sh is None:
            return ""
        c = self.color(color_el(sh), ph) or Clr(0, 0, 0, 0.4)
        dist, ang = iattr(sh, "dist"), _rad(iattr(sh, "dir"))
        blur = iattr(sh, "blurRad")
        return "drop-shadow(%s %s %s %s)" % (self.u(dist * math.cos(ang)), self.u(dist * math.sin(ang)),
                                             self.u(blur / 2.0), c.css())


def _blip_rid(blip):
    if blip is None:
        return None
    for e in blip.iter("{http://schemas.microsoft.com/office/drawing/2016/SVG/main}svgBlip"):
        if e.get(R_EMBED):
            return e.get(R_EMBED)
    return blip.get(R_EMBED)


def _crop_css(src):
    l, t, r, b = (iattr(src, k) / 100000.0 for k in ("l", "t", "r", "b"))
    fw, fh = max(1 - l - r, 1e-4), max(1 - t - b, 1e-4)
    px = l / (l + r) * 100 if abs(l + r) > 1e-6 else 0
    py = t / (t + b) * 100 if abs(t + b) > 1e-6 else 0
    return "background-repeat:no-repeat;background-size:%s%% %s%%;background-position:%s%% %s%%" % (
        num(100 / fw, 3), num(100 / fh, 3), num(px, 3), num(py, 3))


def _shade(c, mode):
    if c is None:
        return Clr(0, 0, 0)
    if mode == "darken":
        return Clr(c.r * .6, c.g * .6, c.b * .6, c.a)
    if mode == "darkenLess":
        return Clr(c.r * .8, c.g * .8, c.b * .8, c.a)
    if mode == "lighten":
        return c.mix(Clr(1, 1, 1, c.a), .4)
    if mode == "lightenLess":
        return c.mix(Clr(1, 1, 1, c.a), .2)
    return c


# ---------------------------------------------------------------------------- images

def _sniff(b):
    if b[:8] == b"\x89PNG\r\n\x1a\n":
        return "png"
    if b[:3] == b"\xff\xd8\xff":
        return "jpeg"
    if b[:6] in (b"GIF87a", b"GIF89a"):
        return "gif"
    if b[:4] == b"RIFF" and b[8:12] == b"WEBP":
        return "webp"
    if b[:2] == b"BM":
        return "bmp"
    if b[:4] in (b"II*\x00", b"MM\x00*"):
        return "tiff"
    if b[:4] == b"II\xbc\x01":
        return "wdp"
    if len(b) > 44 and b[:4] == b"\x01\x00\x00\x00" and b[40:44] == b" EMF":
        return "emf"
    if b[:4] == b"\xd7\xcd\xc6\x9a" or b[:6] in (b"\x01\x00\x09\x00\x00\x03", b"\x02\x00\x09\x00\x00\x03"):
        return "wmf"
    head = b[:512].lstrip()
    if head[:5] == b"<?xml" or head[:4] == b"<svg" or b"<svg" in b[:2048]:
        return "svg"
    if b[:2] == b"\x1f\x8b":
        return "gz"
    return None


class _Img:
    __slots__ = ("cls", "uri", "_blob", "_size")

    def __init__(self, cls, uri, blob, size=None):
        self.cls, self.uri, self._blob, self._size = cls, uri, blob, size

    def size(self):
        if self._size is None:
            try:
                from PIL import Image
                with Image.open(io.BytesIO(self._blob)) as im:
                    self._size = im.size
            except Exception:
                self._size = (96, 96)
        return self._size


class _Images:
    def __init__(self, deck, max_px, recompress_over):
        self.deck = deck
        self.max_px = max_px
        self.recompress_over = recompress_over
        self.map = {}
        self.css = []
        self.bytes = 0

    def get(self, part, rid, duotone=None):
        if not rid or part is None:
            return None
        if duotone is not None:
            return self._duotone(part, rid, duotone)
        try:
            rel = part.rels[rid]
            if rel.is_external:
                return None
            tp = rel.target_part
        except Exception:
            return None
        key = str(tp.partname)
        if key not in self.map:
            try:
                self.map[key] = self._load(tp.blob)
            except Exception as e:  # never let one bad image kill a slide
                self.deck.warn("image %s: %s" % (key, e))
                self.map[key] = None
        return self.map[key]

    def _duotone(self, part, rid, duo):
        try:
            tp = part.rels[rid].target_part
            key = (str(tp.partname), duo[0].css(), duo[1].css())
            if key not in self.map:
                from PIL import Image, ImageOps
                with Image.open(io.BytesIO(tp.blob)) as im:
                    im.load()
                    if self.max_px and max(im.size) > self.max_px:
                        im.thumbnail((self.max_px, self.max_px))
                    rgb = lambda c: tuple(int(round(v * 255)) for v in (c.r, c.g, c.b))  # noqa: E731
                    out = io.BytesIO()
                    ImageOps.colorize(im.convert("L"), rgb(duo[0]), rgb(duo[1])).save(out, "JPEG", quality=85, optimize=True)
                self.map[key] = self._register(out.getvalue(), "image/jpeg", im.size, out.getvalue())
            return self.map[key]
        except Exception as e:
            self.deck.warn("duotone image: %s" % e)
            return self.get(part, rid)

    def _register(self, data, mime, size, blob):
        uri = "data:%s;base64,%s" % (mime, base64.b64encode(data).decode("ascii"))
        cls = "i%d" % len(self.css)
        self.css.append('.%s .%s{background-image:url("%s")}' % (self.deck.uid, cls, uri))
        self.bytes += len(uri)
        return _Img(cls, uri, blob, size)

    def _load(self, blob):
        kind = _sniff(blob)
        if kind == "gz":
            try:
                blob = zlib.decompress(blob, 16 + zlib.MAX_WBITS)
                kind = _sniff(blob)
            except zlib.error:
                return None
        data, mime, size = blob, None, None
        if kind in ("png", "jpeg", "gif", "webp"):
            mime = "image/" + kind
            if kind in ("png", "jpeg") and self.max_px and len(blob) > self.recompress_over:
                data, mime, size = self._shrink(blob, kind)
        elif kind == "svg":
            mime = "image/svg+xml"
        elif kind == "emf":
            svg = emf_to_svg(blob)
            if svg:
                data, mime = svg.encode("utf-8"), "image/svg+xml"
        elif kind == "wmf":
            svg = wmf_to_svg(blob)
            if svg:
                data, mime = svg.encode("utf-8"), "image/svg+xml"
        if mime is None:
            try:  # bmp, tiff, ico, anything Pillow can decode
                from PIL import Image
                with Image.open(io.BytesIO(blob)) as im:
                    im.load()
                    if self.max_px and max(im.size) > self.max_px:
                        im.thumbnail((self.max_px, self.max_px))
                    out = io.BytesIO()
                    im.convert("RGBA" if im.mode in ("RGBA", "LA", "P") else "RGB").save(out, "PNG", optimize=True)
                    data, mime, size = out.getvalue(), "image/png", im.size
            except Exception:
                return None
        if mime == "image/svg+xml" and size is None:
            m = re.search(rb'viewBox="[-\d.]+ [-\d.]+ ([\d.]+) ([\d.]+)"', data[:400])
            size = (float(m.group(1)), float(m.group(2))) if m else (96, 96)
        return self._register(data, mime, size, data)

    def _shrink(self, blob, kind):
        try:
            from PIL import Image
            im = Image.open(io.BytesIO(blob))
            if max(im.size) <= self.max_px and len(blob) < self.recompress_over * 4:
                return blob, "image/" + kind, im.size
            im.load()
            if max(im.size) > self.max_px:
                im.thumbnail((self.max_px, self.max_px), Image.LANCZOS)
            out = io.BytesIO()
            has_alpha = im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info)
            if kind == "jpeg" or not has_alpha:
                im.convert("RGB").save(out, "JPEG", quality=85, optimize=True, progressive=True)
                mime = "image/jpeg"
            else:
                im.save(out, "PNG", optimize=True)
                mime = "image/png"
            if out.tell() >= len(blob):
                return blob, "image/" + kind, None
            return out.getvalue(), mime, im.size
        except Exception:
            return blob, "image/" + kind, None


# ------------------------------------------------------------------- EMF -> SVG

_EMF_STOCK = {
    0: ("brush", (255, 255, 255), 0), 1: ("brush", (192, 192, 192), 0), 2: ("brush", (128, 128, 128), 0),
    3: ("brush", (64, 64, 64), 0), 4: ("brush", (0, 0, 0), 0), 5: ("brush", None, 1),
    6: ("pen", (255, 255, 255), 0, 0), 7: ("pen", (0, 0, 0), 0, 0), 8: ("pen", None, 5, 0),
    18: ("brush", (255, 255, 255), 0), 19: ("pen", (0, 0, 0), 0, 0)}


class _EmfState:
    __slots__ = ("pen", "brush", "font", "text_color", "text_align", "fill_rule", "world", "map_mode",
                 "win_org", "win_ext", "vp_org", "vp_ext", "clip", "arc_dir", "bk_mode")

    def __init__(self):
        self.pen = ("pen", (0, 0, 0), 0, 0)
        self.brush = ("brush", (255, 255, 255), 0)
        self.font = ("font", -12, 400, False, False, False, "Arial", 0)
        self.text_color = (0, 0, 0)
        self.text_align = 0
        self.fill_rule = "evenodd"
        self.world = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
        self.map_mode = 1
        self.win_org = (0, 0)
        self.win_ext = (1, 1)
        self.vp_org = (0, 0)
        self.vp_ext = (1, 1)
        self.clip = None
        self.arc_dir = 1
        self.bk_mode = 2

    def copy(self):
        s = _EmfState()
        for k in self.__slots__:
            setattr(s, k, getattr(self, k))
        return s


def _rgb(t):
    return "#%02x%02x%02x" % t


def _colorref(d, o):
    r, g, b = d[o], d[o + 1], d[o + 2]
    return (r, g, b)


class _EmfRenderer:
    def __init__(self, data):
        self.d = data
        self.s = _EmfState()
        self.stack = []
        self.objs = {}
        self.out = []
        self.defs = []
        self.clip_ids = {}
        self.path = None
        self.cp = (0.0, 0.0)

    # -- coordinate pipeline -------------------------------------------------
    def tp(self, x, y):
        m = self.s.world
        X = x * m[0] + y * m[2] + m[4]
        Y = x * m[1] + y * m[3] + m[5]
        if self.s.map_mode in (7, 8):
            wx, wy = self.s.win_ext
            vx, vy = self.s.vp_ext
            X = (X - self.s.win_org[0]) * (vx / wx if wx else 1) + self.s.vp_org[0]
            Y = (Y - self.s.win_org[1]) * (vy / wy if wy else 1) + self.s.vp_org[1]
        return X, Y

    def scale(self):
        m = self.s.world
        sx = math.hypot(m[0], m[1])
        sy = math.hypot(m[2], m[3])
        if self.s.map_mode in (7, 8):
            wx, wy = self.s.win_ext
            vx, vy = self.s.vp_ext
            sx *= abs(vx / wx) if wx else 1
            sy *= abs(vy / wy) if wy else 1
        return sx, sy

    def P(self, x, y):
        X, Y = self.tp(x, y)
        return "%s %s" % (num(X, 2), num(Y, 2))

    # -- emitting ------------------------------------------------------------
    def _clip_attr(self):
        c = self.s.clip
        if not c:
            return ""
        cid = self.clip_ids.get(c)
        if cid is None:
            cid = "c%d" % len(self.clip_ids)
            self.clip_ids[c] = cid
            self.defs.append('<clipPath id="%s"><rect x="%s" y="%s" width="%s" height="%s"/></clipPath>' % (
                cid, num(c[0], 2), num(c[1], 2), num(max(c[2] - c[0], 0), 2), num(max(c[3] - c[1], 0), 2)))
        return ' clip-path="url(#%s)"' % cid

    def pen_attrs(self):
        p = self.s.pen
        color, style = p[1], p[2] & 0x0F
        if color is None or style == 5:
            return None
        width = p[3]
        sx, sy = self.scale()
        w = max(width * (sx + sy) / 2.0, 1.0) if width else 1.0
        if not (p[2] & 0x10000):  # cosmetic pen
            w = 1.0 if not width else w
        a = 'stroke="%s" stroke-width="%s"' % (_rgb(color), num(w, 2))
        cap = (p[2] >> 8) & 0xF
        join = (p[2] >> 12) & 0xF
        a += ' stroke-linecap="%s" stroke-linejoin="%s"' % ({0: "round", 1: "square", 2: "butt"}.get(cap, "round"),
                                                              {0: "round", 1: "bevel", 2: "miter"}.get(join, "round"))
        dash = {1: (3, 1), 2: (1, 1), 3: (3, 1, 1, 1), 4: (3, 1, 1, 1, 1, 1)}.get(style)
        if dash:
            a += ' stroke-dasharray="%s"' % " ".join(num(v * max(w, 1) * 1.5, 2) for v in dash)
        return a

    def brush_fill(self):
        b = self.s.brush
        if b[1] is None or b[2] == 1:
            return None
        return _rgb(b[1])

    def emit(self, d, fill=True, stroke=True):
        f = self.brush_fill() if fill else None
        st = self.pen_attrs() if stroke else None
        if f is None and st is None:
            return
        self.out.append('<path d="%s" fill="%s"%s%s%s/>' % (
            d, f or "none", ' fill-rule="%s"' % self.s.fill_rule if f else "", " " + st if st else "", self._clip_attr()))

    def add(self, d, fill, stroke):
        if self.path is not None:
            self.path.append(d)
        else:
            self.emit(d, fill, stroke)

    # -- records -------------------------------------------------------------
    def points(self, off, count, small):
        fmt, sz = ("<hh", 4) if small else ("<ii", 8)
        d = self.d
        return [struct.unpack_from(fmt, d, off + i * sz) for i in range(count)]

    def run(self):
        d = self.d
        if len(d) < 88:
            return None
        (typ, size, bl, bt, br, bb, fl, ft, fr, fb, sig, ver, nbytes, nrec, nh, _r, nd, od, npal,
         devx, devy, mmx, mmy) = struct.unpack_from("<II4i4iIIIIHHIII2i2i", d, 0)
        if typ != 1 or sig != 0x464D4520:
            return None
        if mmx and mmy and fr > fl and fb > ft:
            kx, ky = devx / (mmx * 100.0), devy / (mmy * 100.0)
            vb = (fl * kx, ft * ky, (fr - fl) * kx, (fb - ft) * ky)
        else:
            vb = (bl, bt, max(br - bl, 1), max(bb - bt, 1))
        off = size
        n = len(d)
        while off + 8 <= n:
            rt, rs = struct.unpack_from("<II", d, off)
            if rs < 8 or off + rs > n:
                break
            if rt == 14:  # EOF
                break
            try:
                self.record(rt, off, rs)
            except (struct.error, IndexError, ZeroDivisionError, ValueError):
                pass
            off += rs
        if not self.out:
            return None
        return ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="%s %s %s %s" preserveAspectRatio="none">'
                '<defs>%s</defs>%s</svg>') % (num(vb[0], 2), num(vb[1], 2), num(vb[2], 2), num(vb[3], 2),
                                              "".join(self.defs), "".join(self.out))

    def record(self, rt, o, rs):
        d = self.d
        s = self.s
        b = o + 8
        if rt in (85, 86, 87, 88, 89, 2, 3, 4, 5, 6):  # poly shapes
            small = rt >= 85
            cnt = struct.unpack_from("<I", d, b + 16)[0]
            pts = self.points(b + 20, cnt, small)
            if not pts:
                return
            kind = {85: "bez", 2: "bez", 86: "gon", 3: "gon", 87: "line", 4: "line",
                    88: "bezto", 5: "bezto", 89: "lineto", 6: "lineto"}[rt]
            if kind in ("bezto", "lineto"):
                seg = []
                if self.path is None:
                    seg.append("M" + self.P(*self.cp))
                if kind == "bezto":
                    for i in range(0, len(pts) - 2, 3):
                        seg.append("C%s %s %s" % (self.P(*pts[i]), self.P(*pts[i + 1]), self.P(*pts[i + 2])))
                else:
                    seg += ["L" + self.P(*p) for p in pts]
                self.cp = pts[-1]
                self.add("".join(seg), False, True)
                return
            seg = ["M" + self.P(*pts[0])]
            if kind == "bez":
                for i in range(1, len(pts) - 2, 3):
                    seg.append("C%s %s %s" % (self.P(*pts[i]), self.P(*pts[i + 1]), self.P(*pts[i + 2])))
            else:
                seg += ["L" + self.P(*p) for p in pts[1:]]
            if kind == "gon":
                seg.append("Z")
            self.add("".join(seg), kind == "gon", True)
            self.cp = pts[-1]
        elif rt in (90, 91, 7, 8):  # poly-poly
            small = rt >= 90
            npoly, total = struct.unpack_from("<II", d, b + 16)
            counts = struct.unpack_from("<%dI" % npoly, d, b + 24)
            pts = self.points(b + 24 + 4 * npoly, total, small)
            seg, i = [], 0
            gon = rt in (91, 8)
            for c in counts:
                poly = pts[i:i + c]
                i += c
                if poly:
                    seg.append("M" + self.P(*poly[0]) + "".join("L" + self.P(*p) for p in poly[1:]) + ("Z" if gon else ""))
            self.add("".join(seg), gon, True)
        elif rt == 27:  # MOVETOEX
            self.cp = struct.unpack_from("<ii", d, b)
            if self.path is not None:
                self.path.append("M" + self.P(*self.cp))
        elif rt == 54:  # LINETO
            p = struct.unpack_from("<ii", d, b)
            if self.path is not None:
                self.path.append("L" + self.P(*p))
            else:
                self.emit("M%sL%s" % (self.P(*self.cp), self.P(*p)), False, True)
            self.cp = p
        elif rt in (42, 43, 44):  # ELLIPSE, RECTANGLE, ROUNDRECT
            l, t, r, bt = struct.unpack_from("<4i", d, b)
            x0, y0 = self.tp(l, t)
            x1, y1 = self.tp(r, bt)
            x0, x1 = min(x0, x1), max(x0, x1)
            y0, y1 = min(y0, y1), max(y0, y1)
            if rt == 42:
                rx, ry = (x1 - x0) / 2, (y1 - y0) / 2
                cx, cy = x0 + rx, y0 + ry
                dd = "M%s %sA%s %s 0 1 0 %s %sA%s %s 0 1 0 %s %sZ" % tuple(num(v, 2) for v in (
                    cx - rx, cy, rx, ry, cx + rx, cy, rx, ry, cx - rx, cy))
            elif rt == 44:
                cw, ch = struct.unpack_from("<ii", d, b + 16)
                sx, sy = self.scale()
                rx, ry = min(cw * sx / 2, (x1 - x0) / 2), min(ch * sy / 2, (y1 - y0) / 2)
                dd = ("M%s %sH%sA%s %s 0 0 1 %s %sV%sA%s %s 0 0 1 %s %sH%sA%s %s 0 0 1 %s %sV%sA%s %s 0 0 1 %s %sZ" %
                      tuple(num(v, 2) for v in (x0 + rx, y0, x1 - rx, rx, ry, x1, y0 + ry, y1 - ry, rx, ry, x1 - rx, y1,
                                                x0 + rx, rx, ry, x0, y1 - ry, y0 + ry, rx, ry, x0 + rx, y0)))
            else:
                dd = "M%s %sH%sV%sH%sZ" % tuple(num(v, 2) for v in (x0, y0, x1, y1, x0))
            self.add(dd, True, True)
        elif rt in (45, 46, 47, 55):  # ARC, CHORD, PIE, ARCTO
            l, t, r, bt, xs, ys, xe, ye = struct.unpack_from("<8i", d, b)
            x0, y0 = self.tp(l, t)
            x1, y1 = self.tp(r, bt)
            cx, cy, rx, ry = (x0 + x1) / 2, (y0 + y1) / 2, abs(x1 - x0) / 2, abs(y1 - y0) / 2
            if rx <= 0 or ry <= 0:
                return
            sxp, syp = self.tp(xs, ys)
            exp_, eyp = self.tp(xe, ye)
            a1 = math.atan2(syp - cy, sxp - cx)
            a2 = math.atan2(eyp - cy, exp_ - cx)
            p1 = (cx + rx * math.cos(a1), cy + ry * math.sin(a1))
            p2 = (cx + rx * math.cos(a2), cy + ry * math.sin(a2))
            ccw = s.arc_dir == 1
            sweep = (a1 - a2) % (2 * math.pi) if ccw else (a2 - a1) % (2 * math.pi)
            large = 1 if sweep > math.pi else 0
            arc = "A%s %s 0 %d %d %s %s" % (num(rx, 2), num(ry, 2), large, 0 if ccw else 1, num(p2[0], 2), num(p2[1], 2))
            if rt == 55:
                start = ("L" if self.path is not None else "M%s L" % self.P(*self.cp)) + "%s %s" % (num(p1[0], 2), num(p1[1], 2))
                self.add(start + arc, False, True)
            elif rt == 45:
                self.add("M%s %s%s" % (num(p1[0], 2), num(p1[1], 2), arc), False, True)
            elif rt == 46:
                self.add("M%s %s%sZ" % (num(p1[0], 2), num(p1[1], 2), arc), True, True)
            else:
                self.add("M%s %sL%s %s%sZ" % (num(cx, 2), num(cy, 2), num(p1[0], 2), num(p1[1], 2), arc), True, True)
        elif rt == 59:  # BEGINPATH
            self.path = []
        elif rt == 60:  # ENDPATH
            pass
        elif rt == 61:  # CLOSEFIGURE
            if self.path is not None:
                self.path.append("Z")
        elif rt == 68:  # ABORTPATH
            self.path = None
        elif rt in (62, 63, 64):  # FILLPATH, STROKEANDFILLPATH, STROKEPATH
            if self.path:
                dd = "".join(self.path)
                self.path = None
                self.emit(dd, rt in (62, 63), rt in (63, 64))
            self.path = None
        elif rt == 67:  # SELECTCLIPPATH
            self.path = None
        elif rt == 37:  # SELECTOBJECT
            ih = struct.unpack_from("<I", d, b)[0]
            obj = _EMF_STOCK.get(ih & 0x7FFFFFFF) if ih & 0x80000000 else self.objs.get(ih)
            if obj:
                if obj[0] == "pen":
                    s.pen = obj
                elif obj[0] == "brush":
                    s.brush = obj
                elif obj[0] == "font":
                    s.font = obj
        elif rt == 40:  # DELETEOBJECT
            self.objs.pop(struct.unpack_from("<I", d, b)[0], None)
        elif rt == 38:  # CREATEPEN
            ih, style, wx, wy = struct.unpack_from("<IIii", d, b)
            self.objs[ih] = ("pen", None if (style & 0xF) == 5 else _colorref(d, b + 16), style | 0x10000, wx)
        elif rt == 95:  # EXTCREATEPEN
            ih = struct.unpack_from("<I", d, b)[0]
            style, width, bstyle = struct.unpack_from("<III", d, b + 20)
            col = _colorref(d, b + 32)
            self.objs[ih] = ("pen", None if (style & 0xF) == 5 or bstyle == 1 else col, style, width)
        elif rt == 39:  # CREATEBRUSHINDIRECT
            ih, bstyle = struct.unpack_from("<II", d, b)
            self.objs[ih] = ("brush", None if bstyle == 1 else _colorref(d, b + 8), bstyle if bstyle == 1 else 0)
        elif rt in (93, 94):  # CREATEMONOBRUSH / CREATEDIBPATTERNBRUSH -> mid grey
            ih = struct.unpack_from("<I", d, b)[0]
            self.objs[ih] = ("brush", (128, 128, 128), 0)
        elif rt == 82:  # EXTCREATEFONTINDIRECTW
            ih, height, width, esc_, orient, weight = struct.unpack_from("<Iiiiii", d, b)
            italic, underline, strike = d[b + 24], d[b + 25], d[b + 26]
            face = d[b + 32:b + 96].decode("utf-16-le", "ignore").split("\x00")[0]
            self.objs[ih] = ("font", height, weight, bool(italic), bool(underline), bool(strike), face, esc_)
        elif rt == 24:  # SETTEXTCOLOR
            s.text_color = _colorref(d, b)
        elif rt == 22:  # SETTEXTALIGN
            s.text_align = struct.unpack_from("<I", d, b)[0]
        elif rt == 18:
            s.bk_mode = struct.unpack_from("<I", d, b)[0]
        elif rt == 19:  # SETPOLYFILLMODE
            s.fill_rule = "nonzero" if struct.unpack_from("<I", d, b)[0] == 2 else "evenodd"
        elif rt == 17:  # SETMAPMODE
            s.map_mode = struct.unpack_from("<I", d, b)[0]
        elif rt == 9:
            s.win_ext = struct.unpack_from("<ii", d, b)
        elif rt == 10:
            s.win_org = struct.unpack_from("<ii", d, b)
        elif rt == 11:
            s.vp_ext = struct.unpack_from("<ii", d, b)
        elif rt == 12:
            s.vp_org = struct.unpack_from("<ii", d, b)
        elif rt == 57:  # SETARCDIRECTION
            s.arc_dir = struct.unpack_from("<I", d, b)[0]
        elif rt == 35:  # SETWORLDTRANSFORM
            s.world = struct.unpack_from("<6f", d, b)
        elif rt == 36:  # MODIFYWORLDTRANSFORM
            x = struct.unpack_from("<6f", d, b)
            mode = struct.unpack_from("<I", d, b + 24)[0]
            if mode == 1:
                s.world = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
            elif mode == 2:
                s.world = _mat_mul(x, s.world)
            elif mode == 3:
                s.world = _mat_mul(s.world, x)
            elif mode == 4:
                s.world = x
        elif rt == 33:  # SAVEDC
            self.stack.append(s.copy())
        elif rt == 34:  # RESTOREDC
            i = struct.unpack_from("<i", d, b)[0]
            if self.stack:
                idx = len(self.stack) + i if i < 0 else i - 1
                idx = max(0, min(idx, len(self.stack) - 1))
                self.s = self.stack[idx]
                del self.stack[idx:]
        elif rt == 30:  # INTERSECTCLIPRECT
            l, t, r, bt = struct.unpack_from("<4i", d, b)
            x0, y0 = self.tp(l, t)
            x1, y1 = self.tp(r, bt)
            c = (min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))
            if s.clip:
                c = (max(c[0], s.clip[0]), max(c[1], s.clip[1]), min(c[2], s.clip[2]), min(c[3], s.clip[3]))
            s.clip = c
        elif rt == 75:  # EXTSELECTCLIPRGN
            cb, mode = struct.unpack_from("<II", d, b)
            if cb == 0:
                s.clip = None
            else:
                l, t, r, bt = struct.unpack_from("<4i", d, b + 8 + 16)  # RGNDATAHEADER.Bounds (device units)
                s.clip = (float(l), float(t), float(r), float(bt))
        elif rt in (83, 84):  # EXTTEXTOUTA / W
            self.text(b, o, rt == 84)
        elif rt == 81:  # STRETCHDIBITS
            xd, yd, xs, ys, cxs, cys, obmi, cbmi, obits, cbits, usage, rop, cxd, cyd = struct.unpack_from("<4i2i4IIIii", d, b + 16)
            self.bitmap(o, obmi, cbmi, obits, cbits, xd, yd, cxd, cyd)
        elif rt in (76, 77):  # BITBLT / STRETCHBLT
            xd, yd, cxd, cyd, rop = struct.unpack_from("<4iI", d, b + 16)
            obmi, cbmi, obits, cbits = struct.unpack_from("<4I", d, b + 16 + 20 + 8 + 24 + 8)
            if cbmi == 0:
                f = {0x00F00021: self.brush_fill(), 0x00FF0062: "#ffffff", 0x00000042: "#000000"}.get(rop)
                if f:
                    x0, y0 = self.tp(xd, yd)
                    x1, y1 = self.tp(xd + cxd, yd + cyd)
                    self.out.append('<rect x="%s" y="%s" width="%s" height="%s" fill="%s"%s/>' % (
                        num(min(x0, x1), 2), num(min(y0, y1), 2), num(abs(x1 - x0), 2), num(abs(y1 - y0), 2), f, self._clip_attr()))
            else:
                self.bitmap(o, obmi, cbmi, obits, cbits, xd, yd, cxd, cyd)

    def text(self, b, o, wide):
        d = self.d
        s = self.s
        mode, exs, eys = struct.unpack_from("<Iff", d, b + 16)
        rx, ry, nchars, offstr, opts = struct.unpack_from("<iiIII", d, b + 28)
        offdx = struct.unpack_from("<I", d, b + 28 + 20 + 16)[0]
        if not nchars:
            return
        if wide:
            txt = d[o + offstr:o + offstr + 2 * nchars].decode("utf-16-le", "replace")
        else:
            txt = d[o + offstr:o + offstr + nchars].decode("cp1252", "replace")
        font = s.font
        face = font[6] or "Arial"
        txt = map_symbol_text(txt, face)
        if face.lower() in ("symbol", "mt extra", "wingdings"):
            face = "Times New Roman" if face.lower() != "wingdings" else "Segoe UI Symbol"
        txt = clean(txt)
        if not txt.strip():
            return
        sx, sy = self.scale()
        size = abs(font[1]) * sy * (1.0 if font[1] < 0 else 0.85)
        if size <= 0:
            size = 12 * sy
        X, Y = self.tp(rx, ry)
        xs = None
        if offdx:
            adv = struct.unpack_from("<%di" % nchars, d, o + offdx)
            xs, acc = [], 0.0
            for i in range(len(txt)):
                xs.append(X + acc * sx)
                acc += adv[i] if i < len(adv) else 0
            total = acc * sx
        align = s.text_align
        anchor = "start"
        if align & 6 == 6:
            anchor = "middle"
        elif align & 2:
            anchor = "end"
        if xs and anchor != "start":
            shift = total / 2 if anchor == "middle" else total
            xs = [x - shift for x in xs]
            anchor = "start"
        if align & 24 == 24:
            pass  # baseline
        elif align & 8:
            Y -= size * 0.22
        else:
            Y += size * 0.88
        attrs = ['font-family="%s"' % esc(font_stack(face)), 'font-size="%s"' % num(size, 2),
                 'fill="%s"' % _rgb(s.text_color)]
        if font[2] >= 600:
            attrs.append('font-weight="bold"')
        if font[3]:
            attrs.append('font-style="italic"')
        if font[4] or font[5]:
            attrs.append('text-decoration="%s"' % ("underline" if font[4] else "line-through"))
        if anchor != "start":
            attrs.append('text-anchor="%s"' % anchor)
        if font[7]:
            attrs.append('transform="rotate(%s %s %s)"' % (num(-font[7] / 10.0, 2), num(X, 2), num(Y, 2)))
        if xs and len(xs) == len(txt):
            xattr = " ".join(num(x, 2) for x in xs)
        else:
            xattr = num(X, 2)
        self.out.append('<text x="%s" y="%s" xml:space="preserve" %s%s>%s</text>' % (
            xattr, num(Y, 2), " ".join(attrs), self._clip_attr(), esc(txt)))

    def bitmap(self, o, obmi, cbmi, obits, cbits, xd, yd, cxd, cyd):
        d = self.d
        if not cbmi or not cbits:
            return
        bmi = d[o + obmi:o + obmi + cbmi]
        bits = d[o + obits:o + obits + cbits]
        bmp = b"BM" + struct.pack("<IHHI", 14 + len(bmi) + len(bits), 0, 0, 14 + len(bmi)) + bmi + bits
        try:
            from PIL import Image
            with Image.open(io.BytesIO(bmp)) as im:
                im.load()
                out = io.BytesIO()
                im.save(out, "PNG")
        except Exception:
            return
        x0, y0 = self.tp(xd, yd)
        x1, y1 = self.tp(xd + cxd, yd + cyd)
        self.out.append('<image x="%s" y="%s" width="%s" height="%s" preserveAspectRatio="none" href="data:image/png;base64,%s"%s/>' % (
            num(min(x0, x1), 2), num(min(y0, y1), 2), num(abs(x1 - x0), 2), num(abs(y1 - y0), 2),
            base64.b64encode(out.getvalue()).decode("ascii"), self._clip_attr()))


def _mat_mul(a, b):
    """XFORM a * b (apply a first, then b)."""
    return (a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3],
            a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3],
            a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5])


def emf_to_svg(data):
    try:
        return _EmfRenderer(data).run()
    except Exception:
        return None


def wmf_to_svg(data):
    """Minimal WMF support: convert the common drawing records via the EMF renderer's model."""
    try:
        return _WmfRenderer(data).run()
    except Exception:
        return None


class _WmfRenderer(_EmfRenderer):
    def run(self):
        d = self.d
        off = 0
        bbox = None
        unit = 1440
        if d[:4] == b"\xd7\xcd\xc6\x9a":
            l, t, r, bt, unit = struct.unpack_from("<hhhhH", d, 6)
            bbox = (l, t, r, bt)
            off = 22
        hsize = struct.unpack_from("<H", d, off + 2)[0] * 2
        off += hsize
        self.s.map_mode = 8
        self.wobjs = []
        n = len(d)
        while off + 6 <= n:
            size, fn = struct.unpack_from("<IH", d, off)
            size *= 2
            if size < 6 or fn == 0:
                break
            try:
                self.wrec(fn, off + 6)
            except (struct.error, IndexError, ZeroDivisionError):
                pass
            off += size
        if not self.out:
            return None
        if bbox is None:
            ox, oy = self.s.win_org
            ex, ey = self.s.win_ext
            bbox = (ox, oy, ox + ex, oy + ey)
        # map window -> identity viewport, so viewBox is the logical bbox
        return ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="%s %s %s %s" preserveAspectRatio="none">'
                '<defs>%s</defs>%s</svg>') % (bbox[0], bbox[1], max(bbox[2] - bbox[0], 1), max(bbox[3] - bbox[1], 1),
                                              "".join(self.defs), "".join(self.out))

    def tp(self, x, y):
        return float(x), float(y)

    def scale(self):
        return 1.0, 1.0

    def _store(self, obj):
        for i, o in enumerate(self.wobjs):
            if o is None:
                self.wobjs[i] = obj
                return
        self.wobjs.append(obj)

    def wrec(self, fn, p):
        d = self.d
        s = self.s
        h = lambda i: struct.unpack_from("<h", d, p + 2 * i)[0]  # noqa: E731
        if fn == 0x020C:
            s.win_ext = (h(1), h(0))
        elif fn == 0x020B:
            s.win_org = (h(1), h(0))
        elif fn == 0x02FA:  # CREATEPENINDIRECT
            style, w = struct.unpack_from("<Hh", d, p)
            self._store(("pen", None if style & 0xF == 5 else _colorref(d, p + 6), style | 0x10000, max(w, 0)))
        elif fn == 0x02FC:  # CREATEBRUSHINDIRECT
            style = struct.unpack_from("<H", d, p)[0]
            self._store(("brush", None if style == 1 else _colorref(d, p + 2), 1 if style == 1 else 0))
        elif fn == 0x02FB:  # CREATEFONTINDIRECT
            height, width, esc_, orient, weight = struct.unpack_from("<5h", d, p)
            face = d[p + 18:p + 50].split(b"\x00")[0].decode("cp1252", "ignore")
            self._store(("font", height, weight, bool(d[p + 10]), bool(d[p + 11]), bool(d[p + 12]), face, esc_))
        elif fn in (0x00F7, 0x0142, 0x01F9, 0x06FF):  # palette / pattern / region: placeholders
            self._store(("other",))
        elif fn == 0x012D:  # SELECTOBJECT
            i = struct.unpack_from("<H", d, p)[0]
            obj = self.wobjs[i] if i < len(self.wobjs) else None
            if obj and obj[0] == "pen":
                s.pen = obj
            elif obj and obj[0] == "brush":
                s.brush = obj
            elif obj and obj[0] == "font":
                s.font = obj
        elif fn == 0x01F0:  # DELETEOBJECT
            i = struct.unpack_from("<H", d, p)[0]
            if i < len(self.wobjs):
                self.wobjs[i] = None
        elif fn == 0x0209:
            s.text_color = _colorref(d, p)
        elif fn == 0x012E:
            s.text_align = struct.unpack_from("<H", d, p)[0]
        elif fn == 0x0106:
            s.fill_rule = "nonzero" if h(0) == 2 else "evenodd"
        elif fn == 0x0214:  # MOVETO
            self.cp = (h(1), h(0))
        elif fn == 0x0213:  # LINETO
            pnt = (h(1), h(0))
            self.emit("M%sL%s" % (self.P(*self.cp), self.P(*pnt)), False, True)
            self.cp = pnt
        elif fn in (0x0324, 0x0325):  # POLYGON / POLYLINE
            cnt = h(0)
            pts = [(h(1 + 2 * i), h(2 + 2 * i)) for i in range(cnt)]
            if pts:
                self.emit("M" + self.P(*pts[0]) + "".join("L" + self.P(*q) for q in pts[1:]) + ("Z" if fn == 0x0324 else ""),
                          fn == 0x0324, True)
        elif fn == 0x0538:  # POLYPOLYGON
            npoly = h(0)
            counts = [h(1 + i) for i in range(npoly)]
            k = 1 + npoly
            seg = []
            for c in counts:
                pts = [(h(k + 2 * i), h(k + 1 + 2 * i)) for i in range(c)]
                k += 2 * c
                if pts:
                    seg.append("M" + self.P(*pts[0]) + "".join("L" + self.P(*q) for q in pts[1:]) + "Z")
            self.emit("".join(seg), True, True)
        elif fn in (0x041B, 0x0418, 0x061C):  # RECTANGLE, ELLIPSE, ROUNDRECT
            off = 2 if fn == 0x061C else 0
            bt, r, t, l = h(off), h(off + 1), h(off + 2), h(off + 3)
            x0, x1, y0, y1 = min(l, r), max(l, r), min(t, bt), max(t, bt)
            if fn == 0x0418:
                rx, ry = (x1 - x0) / 2, (y1 - y0) / 2
                cx, cy = x0 + rx, y0 + ry
                self.emit("M%s %sA%s %s 0 1 0 %s %sA%s %s 0 1 0 %s %sZ" % tuple(num(v, 2) for v in (
                    cx - rx, cy, rx, ry, cx + rx, cy, rx, ry, cx - rx, cy)), True, True)
            else:
                self.emit("M%s %sH%sV%sH%sZ" % (x0, y0, x1, y1, x0), True, True)
        elif fn in (0x0521, 0x0A32):  # TEXTOUT / EXTTEXTOUT
            if fn == 0x0521:
                cnt = h(0)
                raw = d[p + 2:p + 2 + cnt]
                k = 1 + (cnt + 1) // 2
                y, x = h(k), h(k + 1)
            else:
                y, x, cnt, opts = h(0), h(1), h(2), h(3)
                start = p + 8 + (8 if opts & 0x6 else 0)
                raw = d[start:start + cnt]
            txt = clean(map_symbol_text(raw.decode("cp1252", "replace"), s.font[6]))
            if not txt.strip():
                return
            size = abs(s.font[1]) or 12
            Y = y + (0 if s.text_align & 24 == 24 else size * 0.88)
            anchor = "middle" if s.text_align & 6 == 6 else ("end" if s.text_align & 2 else "start")
            self.out.append('<text x="%s" y="%s" font-family="%s" font-size="%s" fill="%s"%s%s xml:space="preserve">%s</text>' % (
                x, num(Y, 2), esc(font_stack(s.font[6])), num(size, 2), _rgb(s.text_color),
                ' font-weight="bold"' if s.font[2] >= 600 else "",
                ' text-anchor="%s"' % anchor if anchor != "start" else "", esc(txt)))
        elif fn == 0x001E:
            self.stack.append(s.copy())
        elif fn == 0x0127:
            if self.stack:
                self.s = self.stack.pop()


# ------------------------------------------------------------------------------ text

_LINE = 1.2  # PowerPoint "single" line spacing relative to font size
_SAFE_URL = re.compile(r"^(https?://|mailto:)", re.I)
_ALGN = {"l": "left", "ctr": "center", "r": "right", "just": "justify", "dist": "justify"}


def _roman(n):
    out = ""
    for v, s in ((1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"),
                 (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")):
        while n >= v:
            out += s
            n -= v
    return out


def _alpha(n):
    out = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def _autonum(kind, n):
    m = re.match(r"(arabic|alphaLc|alphaUc|romanLc|romanUc)(\w*)", kind or "arabicPeriod")
    base, suf = (m.group(1), m.group(2)) if m else ("arabic", "Period")
    s = {"arabic": str(n), "alphaLc": _alpha(n).lower(), "alphaUc": _alpha(n),
         "romanLc": _roman(n).lower(), "romanUc": _roman(n)}[base]
    if suf == "ParenBoth":
        return "(%s)" % s
    if suf == "ParenR":
        return s + ")"
    if suf == "Plain":
        return s
    return s + "."


def _spacing(el):
    if el is None:
        return None
    p = kid(el, "spcPct")
    if p is not None:
        return ("pct", _int(p.get("val")) / 100000.0)
    p = kid(el, "spcPts")
    if p is not None:
        return ("pts", _int(p.get("val")) / 100.0)
    return None


class _Text:
    """Mixin: text body rendering with full list-style inheritance."""

    def merge_ppr(self, d, el, part):
        if el is None:
            return
        for k in ("algn", "rtl"):
            if el.get(k) is not None:
                d[k] = el.get(k)
        for k in ("marL", "marR", "indent"):
            if el.get(k) is not None:
                d[k] = _int(el.get(k))
        for c in el:
            n = _ln(c)
            if n in ("lnSpc", "spcBef", "spcAft"):
                v = _spacing(c)
                if v:
                    d[n] = v
            elif n == "buNone":
                d["bu"] = None
            elif n == "buChar":
                d["bu"] = ("char", c.get("char") or "•")
            elif n == "buAutoNum":
                d["bu"] = ("auto", c.get("type"), iattr(c, "startAt", 1))
            elif n == "buBlip":
                d["bu"] = ("char", "•")
            elif n == "buFont":
                d["buFont"] = c.get("typeface")
            elif n == "buFontTx":
                d["buFont"] = None
            elif n == "buClr":
                d["buClr"] = self.color(color_el(c))
            elif n == "buClrTx":
                d["buClr"] = None
            elif n == "buSzPct":
                d["buSz"] = _int(c.get("val")) / 100000.0
            elif n == "buSzTx":
                d["buSz"] = None

    def merge_rpr(self, d, el, part):
        if el is None:
            return
        g = el.get
        if g("sz") is not None:
            d["sz"] = _int(g("sz")) / 100.0
        for k in ("b", "i"):
            if g(k) is not None:
                d[k] = g(k) in ("1", "true")
        if g("u") is not None:
            d["u"] = g("u") != "none"
        if g("strike") is not None:
            d["strike"] = g("strike") != "noStrike"
        if g("baseline") is not None:
            d["base"] = _int(g("baseline")) / 1000.0
        if g("cap") is not None:
            d["cap"] = g("cap")
        if g("spc") is not None:
            d["spc"] = _int(g("spc")) / 100.0
        for c in el:
            n = _ln(c)
            if n in ("solidFill", "gradFill", "noFill", "pattFill"):
                holder = etree.Element("x")
                holder.append(copy.deepcopy(c))
                f = self.fill(holder, part)
                if f is not None:
                    d["fill"] = f
            elif n in ("latin", "ea", "sym"):
                if c.get("typeface"):
                    d[n] = c.get("typeface")
            elif n == "highlight":
                d["hl"] = self.color(color_el(c))
            elif n == "hlinkClick":
                d["link"] = c.get(R_ID)

    def level_style(self, sources, lvl, part):
        key = (tuple(id(s) for s in sources), lvl, id(self.clrmap), id(self.theme))
        hit = self.deck.lvl_cache.get(key)
        if hit is None:
            p, r = {}, {}
            tag = "lvl%dpPr" % (lvl + 1)
            for s in sources:
                e = kid(s, tag)
                if e is not None:
                    self.merge_ppr(p, e, part)
                    self.merge_rpr(r, kid(e, "defRPr"), part)
            hit = (p, r, sources)  # keep sources alive so ids stay unique
            self.deck.lvl_cache[key] = hit
        return dict(hit[0]), dict(hit[1])

    def run_css(self, r, scale=1.0):
        """Run property dict -> (inheritable css dict, non-inheritable css string)."""
        c = {}
        sz = r.get("sz", 18.0) * scale
        base = r.get("base")
        if base:
            sz_css = sz * 0.66
        else:
            sz_css = sz
        c["font-size"] = self.u(sz_css * EMU_PT)
        c["font-weight"] = "700" if r.get("b") else "400"
        c["font-style"] = "italic" if r.get("i") else "normal"
        col = self.fill_color(r.get("fill")) or self.scheme("tx1")
        f = r.get("fill")
        if f and f[0] == "none":
            col = Clr(0, 0, 0, 0)
        c["color"] = col.css()
        face = self.theme.font(r.get("latin") or "+mn-lt")
        c["font-family"] = font_stack(face, self.theme.font(r.get("ea")) if r.get("ea") else None)
        if r.get("spc"):
            c["letter-spacing"] = self.u(r["spc"] * EMU_PT)
        if r.get("cap") == "all":
            c["text-transform"] = "uppercase"
        elif r.get("cap") == "small":
            c["font-variant"] = "small-caps"
        extra = []
        deco = [k for k, on in (("underline", r.get("u")), ("line-through", r.get("strike"))) if on]
        if deco:
            extra.append("text-decoration:" + " ".join(deco))
        if base:
            extra.append("vertical-align:%s%%;line-height:0" % num(base, 1))
        if r.get("hl") is not None:
            extra.append("background:" + r["hl"].css())
        return c, ";".join(extra)

    def _link(self, rid, part):
        try:
            rel = part.rels[rid]
            if rel.is_external and _SAFE_URL.match(rel.target_ref or ""):
                return rel.target_ref
        except Exception:
            pass
        return None

    def paragraphs(self, txBody, part, inherited, extra=None, own=None, font_scale=1.0, ln_reduce=0.0,
                   nowrap=False, default_algn=None):
        """Render the a:p children of txBody -> html string (empty if no visible text)."""
        out = []
        counters = {}
        has_text = False
        first = True
        for p in kids(txBody, "p"):
            pPr = kid(p, "pPr")
            lvl = min(iattr(pPr, "lvl"), 8)
            pp, base = self.level_style(inherited, lvl, part)
            if extra:
                base.update(extra)
            if own is not None:
                e = kid(own, "lvl%dpPr" % (lvl + 1))
                if e is not None:
                    self.merge_ppr(pp, e, part)
                    self.merge_rpr(base, kid(e, "defRPr"), part)
            self.merge_ppr(pp, pPr, part)
            # runs
            runs = []
            for c in p:
                n = _ln(c)
                if n in ("r", "fld"):
                    r = dict(base)
                    self.merge_rpr(r, kid(c, "rPr"), part)
                    t = kid(c, "t")
                    text = t.text or "" if t is not None else ""
                    if n == "fld" and (c.get("type") or "").startswith("slidenum") and self.slide_no:
                        text = str(self.slide_no)
                    if text:
                        runs.append((r, text))
                elif n == "br":
                    r = dict(base)
                    self.merge_rpr(r, kid(c, "rPr"), part)
                    runs.append((r, None))
                elif n == "m" or n == "AlternateContent":
                    txt = "".join(c.itertext()).strip()
                    if txt:
                        runs.append((dict(base), txt))
            if runs:
                pr = runs[0][0]
            else:
                pr = dict(base)
                self.merge_rpr(pr, kid(p, "endParaRPr"), part)
            pcss, _ = self.run_css(pr, font_scale)
            st = ["%s:%s" % kv for kv in pcss.items()]
            algn = pp.get("algn") or default_algn
            if algn and algn != "l":
                st.append("text-align:" + _ALGN.get(algn, "left"))
            if pp.get("rtl") == "1":
                st.append("direction:rtl")
            ls = pp.get("lnSpc")
            if ls and ls[0] == "pts":
                st.append("line-height:" + self.u(ls[1] * EMU_PT * font_scale))
            else:
                pct = (ls[1] if ls else 1.0) - ln_reduce
                if abs(pct - 1.0) > 1e-3:
                    st.append("line-height:%s" % num(max(pct, 0.1) * _LINE, 3))
            psz = pr.get("sz", 18.0) * font_scale
            for key, prop in (("spcBef", "margin-top"), ("spcAft", "margin-bottom")):
                sp = pp.get(key)
                if sp and not (first and key == "spcBef"):
                    v = sp[1] * psz * _LINE if sp[0] == "pct" else sp[1] * font_scale
                    if v > 0.01:
                        st.append("%s:%s" % (prop, self.u(v * EMU_PT)))
            marL, indent = pp.get("marL", 0), pp.get("indent", 0)
            bu = pp.get("bu")
            bullet = ""
            if bu and runs:
                if bu[0] == "auto":
                    k = (lvl, bu[1])
                    counters[k] = counters.get(k, bu[2] - 1) + 1
                    for kk in [x for x in counters if x[0] > lvl]:
                        del counters[kk]
                    ch = _autonum(bu[1], counters[k])
                    bfont = None
                else:
                    counters.clear() if lvl == 0 else None
                    bfont = pp.get("buFont")
                    ch = map_symbol_text(bu[1], bfont)
                    if bfont and bfont.lower() in ("symbol",) or (bfont or "").lower().startswith("wingdings"):
                        bfont = None
                bst = ["display:inline-block", "text-indent:0", "text-align:left"]
                if indent < 0:
                    bst.append("min-width:" + self.u(-indent))
                else:
                    bst.append("margin-right:.3em")
                if pp.get("buClr") is not None:
                    bst.append("color:" + pp["buClr"].css())
                if pp.get("buSz"):
                    bst.append("font-size:%s%%" % num(pp["buSz"] * 100, 1))
                if bfont:
                    bst.append("font-family:" + font_stack(self.theme.font(bfont)))
                bst.append("font-weight:400;font-style:normal;text-decoration:none")
                bullet = '<span style="%s">%s</span>' % (esc(";".join(bst)), _html.escape(clean(ch), quote=False))
            else:
                if not bu:
                    counters.clear()
            if marL:
                st.append("margin-left:" + self.u(marL))
            if indent and (bullet or indent > 0 or marL):
                st.append("text-indent:" + self.u(indent if (bullet or indent > 0) else max(indent, -marL)))
            if pp.get("marR"):
                st.append("margin-right:" + self.u(pp["marR"]))
            body = [bullet]
            for r, text in runs:
                if text is None:
                    body.append("<br>")
                    continue
                has_text = has_text or bool(text.strip())
                face = self.theme.font(r.get("latin") or "")
                if r.get("sym") and any("\uf000" <= ch <= "\uf0ff" for ch in text):
                    sym = r["sym"]
                    text = "".join(map_symbol_text(ch, sym) if "\uf000" <= ch <= "\uf0ff" else ch for ch in text)
                text = clean(map_symbol_text(text, face) if face else text)
                if face and (face.lower() == "symbol" or face.lower().startswith("wingdings")):
                    r = dict(r, latin="Segoe UI Symbol")
                rcss, rextra = self.run_css(r, font_scale)
                diff = ["%s:%s" % (k, v) for k, v in rcss.items() if pcss.get(k) != v]
                href = self._link(r["link"], part) if r.get("link") else None
                if href:
                    if "fill" not in r or True:
                        diff.append("color:%s;text-decoration:underline" % self.scheme("hlink").css())
                if rextra:
                    diff.append(rextra)
                t = _html.escape(text, quote=False).replace("\x0b", "<br>").replace("\t", '<span class="tab">\t</span>')
                if href:
                    t = '<a href="%s" target="_blank" rel="noopener noreferrer"%s>%s</a>' % (
                        _html.escape(href.replace('"', "%22"), quote=False), ' style="%s"' % esc(";".join(diff)) if diff else "", t)
                elif diff:
                    t = '<span style="%s">%s</span>' % (esc(";".join(diff)), t)
                body.append(t)
            if not runs:
                body.append("&#8203;")
            out.append('<p style="%s">%s</p>' % (esc(";".join(st)), "".join(body)))
            first = False
        return "".join(out) if has_text else ""

    def body_props(self, bodyPrs):
        """bodyPrs: highest priority first. -> merged dict."""
        d = {"lIns": 91440, "tIns": 45720, "rIns": 91440, "bIns": 45720, "anchor": "t", "wrap": "square",
             "vert": "horz", "anchorCtr": False, "scale": 1.0, "reduce": 0.0}
        for bp in reversed([b for b in bodyPrs if b is not None]):
            for k in ("lIns", "tIns", "rIns", "bIns"):
                if bp.get(k) is not None:
                    d[k] = _int(bp.get(k))
            for k in ("anchor", "wrap", "vert"):
                if bp.get(k) is not None:
                    d[k] = bp.get(k)
            if bp.get("anchorCtr") is not None:
                d["anchorCtr"] = bp.get("anchorCtr") in ("1", "true")
        top = bodyPrs[0] if bodyPrs and bodyPrs[0] is not None else None
        na = kid(top, "normAutofit")
        if na is not None:
            d["scale"] = iattr(na, "fontScale", 100000) / 100000.0
            d["reduce"] = iattr(na, "lnSpcReduction", 0) / 100000.0
        return d

    def text_box(self, txBody, part, rect, bodyPrs, inherited, extra=None, own=None, default_algn=None):
        """Full text box for a shape. rect = (l, t, r, b) EMU within the shape box."""
        if txBody is None:
            return ""
        bp = self.body_props(bodyPrs)
        nowrap = bp["wrap"] == "none"
        inner = self.paragraphs(txBody, part, inherited, extra, own, bp["scale"], bp["reduce"], nowrap, default_algn)
        if not inner:
            return ""
        l, t, r, b = rect
        st = ["left:%s;top:%s;width:%s;height:%s" % (self.u(l), self.u(t), self.u(max(r - l, 0)), self.u(max(b - t, 0))),
              "padding:%s %s %s %s" % (self.u(bp["tIns"]), self.u(bp["rIns"]), self.u(bp["bIns"]), self.u(bp["lIns"]))]
        jc = {"t": "flex-start", "ctr": "center", "b": "flex-end"}.get(bp["anchor"], "flex-start")
        if jc != "flex-start":
            st.append("justify-content:" + jc)
        cls = "t"
        if nowrap:
            cls += " nw"
            algn = None
            first = path(txBody, "p", "pPr")
            if first is not None:
                algn = first.get("algn")
            st.append("align-items:" + {"ctr": "center", "r": "flex-end"}.get(algn or default_algn, "flex-start"))
        elif bp["anchorCtr"]:
            st.append("align-items:center")
        if bp["vert"] in ("vert", "eaVert", "wordArtVert", "wordArtVertRtl", "mongolianVert"):
            st.append("writing-mode:vertical-rl")
        elif bp["vert"] == "vert270":
            st.append("writing-mode:vertical-rl;transform:rotate(180deg)")
        return '<div class="%s" style="%s">%s</div>' % (cls, esc(";".join(st)), inner)


# ---------------------------------------------------------------------------- shapes

_TITLE_PH = ("title", "ctrTitle")
_OTHER_PH = ("dt", "ftr", "sldNum")
IDENT = (1.0, 1.0, 0.0, 0.0)


def _xfrm(el):
    """a:xfrm / p:xfrm -> (x, y, cx, cy, rot_deg, flipH, flipV) or None."""
    if el is None:
        return None
    off, ext = kid(el, "off"), kid(el, "ext")
    if off is None or ext is None:
        return None
    return (iattr(off, "x"), iattr(off, "y"), iattr(ext, "cx"), iattr(ext, "cy"),
            iattr(el, "rot") / 60000.0, battr(el, "flipH"), battr(el, "flipV"))


def _ph(sp):
    for nv in sp:
        if _ln(nv).startswith("nv"):
            return path(nv, "nvPr", "ph")
    return None


def _hidden(sp):
    for nv in sp:
        if _ln(nv).startswith("nv"):
            return battr(kid(nv, "cNvPr"), "hidden")
    return False


def _ph_index(tree):
    """spTree -> ([(type, idx, sp)], ...) for its placeholder shapes."""
    out = []
    if tree is None:
        return out
    for sp in tree:
        if _ln(sp) in ("sp", "pic", "graphicFrame"):
            ph = _ph(sp)
            if ph is not None:
                out.append((ph.get("type") or "body", ph.get("idx"), sp))
    return out


def _same_type(a, b):
    if a == b:
        return True
    return a in _TITLE_PH and b in _TITLE_PH


class _Shapes:
    """Mixin: shape tree rendering."""

    # placeholder inheritance ---------------------------------------------------
    def find_layout_ph(self, ph):
        idx = self.deck.ph_cache.get(id(self.layout_el))
        if idx is None:
            idx = _ph_index(path(self.layout_el, "cSld", "spTree"))
            self.deck.ph_cache[id(self.layout_el)] = idx
        typ, i = ph.get("type") or "body", ph.get("idx")
        if i is not None:
            for t, j, sp in idx:
                if j == i:
                    return sp
        for t, j, sp in idx:
            if _same_type(t, typ) and (i is None or typ != "body"):
                return sp
        return None

    def find_master_ph(self, typ):
        idx = self.deck.ph_cache.get(id(self.master_el))
        if idx is None:
            idx = _ph_index(path(self.master_el, "cSld", "spTree"))
            self.deck.ph_cache[id(self.master_el)] = idx
        if typ in _TITLE_PH:
            want = "title"
        elif typ in _OTHER_PH:
            want = typ
        else:
            want = "body"
        for t, j, sp in idx:
            if t == want:
                return sp
        return None

    def inherit(self, sp, layer):
        """-> (layout_sp, master_sp, ph_type) for a slide/layout placeholder shape."""
        ph = _ph(sp)
        if ph is None:
            return None, None, None
        typ = ph.get("type") or "body"
        lay = None
        if layer == "slide":
            lay = self.find_layout_ph(ph)
            if lay is not None:
                lph = _ph(lay)
                if ph.get("type") is None and lph is not None and lph.get("type"):
                    typ = lph.get("type")
        mas = self.find_master_ph(typ) if layer != "master" else None
        return lay, mas, typ

    def text_sources(self, typ, lay, mas):
        tx = kid(self.master_el, "txStyles")
        if typ is None:
            return [s for s in (self.deck.default_text, kid(tx, "otherStyle")) if s is not None]
        if typ in _TITLE_PH:
            base = kid(tx, "titleStyle")
        elif typ in _OTHER_PH:
            base = kid(tx, "otherStyle")
        else:
            base = kid(tx, "bodyStyle")
        src = [base, path(mas, "txBody", "lstStyle"), path(lay, "txBody", "lstStyle")]
        return [s for s in src if s is not None]

    # tree ----------------------------------------------------------------------
    def tree(self, parent, part, tf=IDENT, layer="slide", grp_fill=None):
        out = []
        for el in parent:
            n = _ln(el)
            try:
                if n == "AlternateContent":
                    branch = kid(el, "Fallback")
                    if branch is None or len(branch) == 0:
                        branch = kid(el, "Choice")
                    if branch is not None:
                        out.append(self.tree(branch, part, tf, layer, grp_fill))
                    continue
                if n not in ("sp", "pic", "cxnSp", "grpSp", "graphicFrame") or _hidden(el):
                    continue
                if layer != "slide" and n != "grpSp" and _ph(el) is not None:
                    continue
                if n == "grpSp":
                    out.append(self.group(el, part, tf, layer, grp_fill))
                elif n == "pic":
                    out.append(self.picture(el, part, tf, layer))
                elif n == "graphicFrame":
                    out.append(self.frame(el, part, tf, layer))
                else:
                    out.append(self.shape(el, part, tf, layer, grp_fill))
            except Exception as e:  # one broken shape must not take the slide down
                self.deck.warn("slide %s: %s skipped (%s: %s)" % (self.slide_no, n, type(e).__name__, e))
        return "".join(out)

    def box(self, xf, tf, extra_style="", cls="s"):
        """Positioned wrapper opening tag + rendered (w, h)."""
        x, y, cx, cy, rot, fh, fv = xf
        sx, sy, ox, oy = tf
        w, h = cx * sx, cy * sy
        st = "left:%s;top:%s;width:%s;height:%s" % (self.u((x - ox) * sx), self.u((y - oy) * sy), self.u(w), self.u(h))
        if rot:
            st += ";transform:rotate(%sdeg)" % num(rot, 2)
        if extra_style:
            st += ";" + extra_style
        return '<div class="%s" style="%s">' % (cls, esc(st)), w, h

    def group(self, el, part, tf, layer, grp_fill):
        pr = kid(el, "grpSpPr")
        xe = kid(pr, "xfrm")
        xf = _xfrm(xe)
        if xf is None:
            return self.tree(el, part, tf, layer, grp_fill)
        co, ce = kid(xe, "chOff"), kid(xe, "chExt")
        cw, ch = iattr(ce, "cx", xf[2]) or 1, iattr(ce, "cy", xf[3]) or 1
        sx, sy = tf[0], tf[1]
        ctf = (sx * (xf[2] / cw if cw else 1.0), sy * (xf[3] / ch if ch else 1.0), iattr(co, "x", xf[0]), iattr(co, "y", xf[1]))
        flip = ""
        if xf[5] or xf[6]:
            flip = "transform:rotate(%sdeg) scale(%d,%d)" % (num(xf[4], 2), -1 if xf[5] else 1, -1 if xf[6] else 1)
            xf = xf[:4] + (0.0, False, False)
        gf = self.fill(pr, part, grp=grp_fill) or grp_fill
        head, _, _ = self.box(xf, tf, flip)
        return head + self.tree(el, part, ctf, layer, gf) + "</div>"

    def geometry_svg(self, geom, w, h, fill, line, flip, shadow, mode_fill=True):
        """Geometry + paint -> svg string ('' when nothing visible)."""
        defs = []
        stroke, lw = self.stroke_attrs(line, defs)
        has_fill = mode_fill and fill and fill[0] not in ("none", "img")
        if not has_fill and not lw:
            return ""
        pad = lw * 6 + 12700
        body = []
        for d, fmode, strk in geom.paths:
            if not d:
                continue
            paint = self.svg_paint(fill, defs, fmode) if has_fill else "none"
            s = stroke if (strk and lw) else 'stroke="none"'
            if paint == "none" and s == 'stroke="none"':
                continue
            body.append('<path d="%s" fill="%s" %s/>' % (d, paint, s))
        if not body:
            return ""
        st = "left:%s;top:%s;width:%s;height:%s" % (self.u(-pad), self.u(-pad), self.u(w + 2 * pad), self.u(h + 2 * pad))
        if flip[0] or flip[1]:
            st += ";transform:scale(%d,%d)" % (-1 if flip[0] else 1, -1 if flip[1] else 1)
        if shadow:
            st += ";filter:" + shadow
        return '<svg class="g" viewBox="%s %s %s %s" preserveAspectRatio="none" style="%s">%s%s</svg>' % (
            num(-pad, 0), num(-pad, 0), num(w + 2 * pad, 0), num(h + 2 * pad, 0), esc(st),
            "<defs>%s</defs>" % "".join(defs) if defs else "", "".join(body))

    def clip_css(self, geom, w, h, spPr_list):
        """clip-path for image layers; '' for plain rectangles."""
        prst = None
        for spPr in spPr_list:
            if kid(spPr, "custGeom") is not None:
                prst = "cust"
                break
            g = kid(spPr, "prstGeom")
            if g is not None:
                prst = g.get("prst")
                break
        if prst in (None, "rect"):
            return ""
        if prst == "ellipse":
            return "border-radius:50%"
        poly = geom.polys[0] if geom.polys else []
        if len(poly) < 3:
            return ""
        return "clip-path:polygon(%s)" % ",".join(
            "%s%% %s%%" % (num(px / w * 100, 2), num(py / h * 100, 2)) for px, py in poly)

    def shape(self, sp, part, tf, layer, grp_fill):
        spPr, style, txBody = kid(sp, "spPr"), kid(sp, "style"), kid(sp, "txBody")
        lay, mas, typ = self.inherit(sp, layer)
        chain = [(spPr, part), (kid(lay, "spPr"), self.layout_part), (kid(mas, "spPr"), self.master_part)]
        xf = None
        for pr, _p in chain:
            xf = _xfrm(kid(pr, "xfrm"))
            if xf is not None:
                break
        if xf is None:
            return ""
        spPrs = [c[0] for c in chain if c[0] is not None]
        fill = None
        for pr, p in chain:
            fill = self.fill(pr, p, grp=grp_fill)
            if fill is not None:
                break
        if fill is None and style is not None:
            fill = self.style_fill(style, part)
        line = self.style_line(style, part) if style is not None else {}
        for pr, p in reversed(chain):
            line.update(self.line_props(kid(pr, "ln"), p))
        head, w, h = self.box(xf, tf)
        visible = (fill and fill[0] != "none") or (line.get("fill") and line["fill"][0] != "none")
        geom = None
        parts = []
        if visible:
            geom = build_geometry(spPrs, w, h)
            effects, eph = self.effect_list(spPrs, style)
            shadow = self.shadow_css(effects, eph)
            if fill and fill[0] == "img":
                css, cls = self.fill_css(fill)
                clip = self.clip_css(geom, max(w, 1), max(h, 1), spPrs)
                parts.append('<div class="im %s" style="%s"></div>' % (cls, esc(css + (";" + clip if clip else ""))))
            parts.append(self.geometry_svg(geom, w, h, fill, line, (xf[5], xf[6]), shadow))
        if txBody is not None:
            if geom is None:
                has_geom = any(kid(pr, "custGeom") is not None or
                               (kid(pr, "prstGeom") is not None and kid(pr, "prstGeom").get("prst") != "rect") for pr in spPrs)
                rect = build_geometry(spPrs, w, h).rect if has_geom else (0.0, 0.0, w, h)
            else:
                rect = geom.rect
            extra = None
            fr = kid(style, "fontRef")
            c = self.color(color_el(fr)) if fr is not None else None
            if c is not None:
                extra = {"fill": ("solid", c)}
            bodyPrs = [kid(txBody, "bodyPr"), path(lay, "txBody", "bodyPr"), path(mas, "txBody", "bodyPr")]
            parts.append(self.text_box(txBody, part, rect, bodyPrs, self.text_sources(typ, lay, mas),
                                       extra, kid(txBody, "lstStyle")))
        inner = "".join(parts)
        return head + inner + "</div>" if inner else ""

    def picture(self, pic, part, tf, layer):
        spPr = kid(pic, "spPr")
        lay, mas, typ = self.inherit(pic, layer)
        xf = _xfrm(kid(spPr, "xfrm")) or _xfrm(path(lay, "spPr", "xfrm")) or _xfrm(path(mas, "spPr", "xfrm"))
        if xf is None:
            return ""
        bf = kid(pic, "blipFill")
        blip = kid(bf, "blip")
        img = self.deck.images.get(part, _blip_rid(blip)) if blip is not None else None
        head, w, h = self.box(xf, tf)
        spPrs = [p for p in (spPr, kid(lay, "spPr")) if p is not None]
        geom = build_geometry(spPrs, w, h)
        parts = []
        st = []
        flip = (xf[5], xf[6])
        if flip[0] or flip[1]:
            st.append("transform:scale(%d,%d)" % (-1 if flip[0] else 1, -1 if flip[1] else 1))
        clip = self.clip_css(geom, max(w, 1), max(h, 1), spPrs)
        if clip:
            st.append(clip)
        effects, eph = self.effect_list(spPrs, kid(pic, "style"))
        shadow = self.shadow_css(effects, eph)
        if img is not None:
            src = kid(bf, "srcRect")
            if src is not None and any(src.get(k) for k in ("l", "t", "r", "b")):
                st.append(_crop_css(src))
            alpha = iattr(kid(blip, "alphaModFix"), "amt", 100000) / 100000.0
            if alpha < 0.999:
                st.append("opacity:%s" % num(alpha, 3))
            if shadow and not clip:
                st.append("filter:" + shadow)
            parts.append('<div class="im %s" style="%s"></div>' % (img.cls, esc(";".join(st))))
        else:
            parts.append('<div class="im missing"></div>')
        line = self.style_line(kid(pic, "style"), part) if kid(pic, "style") is not None else {}
        line.update(self.line_props(kid(spPr, "ln"), part))
        parts.append(self.geometry_svg(geom, w, h, None, line, flip, "", mode_fill=False))
        return head + "".join(parts) + "</div>"

    def frame(self, fr, part, tf, layer):
        xf = _xfrm(kid(fr, "xfrm"))
        if xf is None and layer == "slide":
            lay, mas, _t = self.inherit(fr, layer)
            xf = _xfrm(path(lay, "spPr", "xfrm")) or _xfrm(path(mas, "spPr", "xfrm"))
        if xf is None:
            return ""
        data = path(fr, "graphic", "graphicData")
        if data is None:
            return ""
        head, w, h = self.box(xf, tf)
        tbl = kid(data, "tbl")
        if tbl is not None:
            return head + self.table(tbl, part, w, h, tf) + "</div>"
        for el in data.iter():
            n = _ln(el)
            if n == "chart" and el.get(R_ID):
                return head + self.chart(part, el.get(R_ID), w, h) + "</div>"
            if n == "relIds" and el.get(R_DM):
                inner = self.smartart(part, el, w, h)
                return head + inner + "</div>" if inner else ""
            if n == "pic":  # OLE object / embedded media preview image
                blip = path(el, "blipFill", "blip")
                img = self.deck.images.get(part, _blip_rid(blip)) if blip is not None else None
                if img is not None:
                    return head + '<div class="im %s"></div></div>' % img.cls
        return ""

    def smartart(self, part, rel_el, w, h):
        """Render SmartArt from its cached drawing part (dsp:drawing)."""
        drawing = None
        try:
            dm = part.rels[rel_el.get(R_DM)].target_part
            root = etree.fromstring(dm.blob)
            for e in root.iter():
                if _ln(e) == "dataModelExt" and e.get("relId"):
                    drawing = part.rels[e.get("relId")].target_part
                    break
        except Exception:
            drawing = None
        if drawing is None:
            for rel in part.rels.values():
                if rel.reltype.endswith("/diagramDrawing") and not rel.is_external:
                    drawing = rel.target_part
                    break
        if drawing is None:
            return ""
        root = etree.fromstring(drawing.blob)
        tree = kid(root, "spTree")
        if tree is None:
            return ""
        out = []
        for sp in kids(tree, "sp"):
            # dsp:sp keeps its text rectangle in txXfrm; lift text into a sibling box
            out.append(self.shape(sp, drawing, IDENT, "smartart", None))
        return "".join(out)


# ---------------------------------------------------------------------------- tables

_ROW_REGIONS = ("firstRow", "lastRow", "band1H", "band2H")
_COL_REGIONS = ("firstCol", "lastCol", "band1V", "band2V")


class _Tables:
    def _regions(self, tblPr, r, c, nr, nc):
        """Applicable table-style regions for a cell, lowest priority first."""
        f = lambda k: battr(tblPr, k)  # noqa: E731
        out = ["wholeTbl"]
        first_row, last_row = f("firstRow") and r == 0, f("lastRow") and r == nr - 1
        first_col, last_col = f("firstCol") and c == 0, f("lastCol") and c == nc - 1
        if f("bandRow") and not first_row and not last_row:
            out.append("band1H" if (r - (1 if f("firstRow") else 0)) % 2 == 0 else "band2H")
        if f("bandCol") and not first_col and not last_col:
            out.append("band1V" if (c - (1 if f("firstCol") else 0)) % 2 == 0 else "band2V")
        if last_col:
            out.append("lastCol")
        if first_col:
            out.append("firstCol")
        if last_row:
            out.append("lastRow")
        if first_row:
            out.append("firstRow")
        return out

    def _fallback_style(self, region):
        """Medium Style 2 - Accent 1 (PowerPoint's default), used when the style part is missing."""
        acc, lt = self.scheme("accent1"), self.scheme("lt1")
        edge = {"w": 12700, "fill": ("solid", lt)}
        bd = {k: edge for k in ("left", "right", "top", "bottom", "insideH", "insideV")}
        if region == "wholeTbl":
            return {"fill": ("solid", acc.mix(lt, 0.8)), "bd": bd, "color": self.scheme("dk1")}
        if region == "band1H" or region == "band1V":
            return {"fill": ("solid", acc.mix(lt, 0.6))}
        if region in ("firstRow", "lastRow", "firstCol", "lastCol"):
            return {"fill": ("solid", acc), "color": lt, "b": True}
        return {}

    def _style_region(self, style, region, part):
        key = (id(style), region, id(self.clrmap))
        hit = self.deck.tbl_cache.get(key)
        if hit is not None:
            return hit
        d = {}
        if style is None:
            d = self._fallback_style(region)
        else:
            el = kid(style, region)
            if el is not None:
                tcs = kid(el, "tcStyle")
                fl = kid(tcs, "fill")
                if fl is not None:
                    f = self.fill(fl, part)
                    if f is not None:
                        d["fill"] = f
                fr = kid(tcs, "fillRef")
                if fr is not None and "fill" not in d:
                    f = self.style_fill(tcs, part)
                    if f is not None:
                        d["fill"] = f
                bd = {}
                for b in (kid(tcs, "tcBdr") if kid(tcs, "tcBdr") is not None else []):
                    ln = kid(b, "ln")
                    if ln is not None:
                        bd[_ln(b)] = self.line_props(ln, part)
                    elif kid(b, "lnRef") is not None:
                        ref = kid(b, "lnRef")
                        i = iattr(ref, "idx")
                        if 0 < i <= len(self.theme.lines):
                            bd[_ln(b)] = self.line_props(self.theme.lines[i - 1], part, self.color(color_el(ref)))
                if bd:
                    d["bd"] = bd
                tx = kid(el, "tcTxStyle")
                if tx is not None:
                    if tx.get("b") is not None:
                        d["b"] = tx.get("b") == "on"
                    if tx.get("i") is not None:
                        d["i"] = tx.get("i") == "on"
                    c = self.color(color_el(tx))
                    if c is None and kid(tx, "fontRef") is not None:
                        c = self.color(color_el(kid(tx, "fontRef")))
                    if c is not None:
                        d["color"] = c
        self.deck.tbl_cache[key] = d
        return d

    def _border(self, ln):
        f = ln.get("fill") if ln else None
        if not f or f[0] == "none":
            return None
        c = self.fill_color(f) or Clr(0, 0, 0)
        return "%s %s %s" % (self.u(max(ln.get("w", 12700), 6350)), "dashed" if ln.get("dash") else "solid", c.css())

    def table(self, tbl, part, w, h, tf):
        tblPr = kid(tbl, "tblPr")
        sid = kid(tblPr, "tableStyleId")
        has_style = sid is not None and (sid.text or "").strip()
        style = self.deck.table_styles.get((sid.text or "").strip()) if has_style else None
        cols = [iattr(g, "w") for g in kids(kid(tbl, "tblGrid"), "gridCol")]
        total = float(sum(cols)) or 1.0
        rows = kids(tbl, "tr")
        nr, nc = len(rows), len(cols)
        tbl_fill = self.fill(tblPr, part)
        out = ['<table class="tb"%s><colgroup>' % (' style="%s"' % esc(self.fill_css(tbl_fill)[0]) if tbl_fill else "")]
        out += ['<col style="width:%s%%">' % num(c / total * 100, 3) for c in cols]
        out.append("</colgroup>")
        sources = [s for s in (self.deck.default_text,) if s is not None]
        for r, tr in enumerate(rows):
            out.append('<tr style="height:%s">' % self.u(iattr(tr, "h") * tf[1]))
            c = 0
            for tc in kids(tr, "tc"):
                span, rspan = iattr(tc, "gridSpan", 1), iattr(tc, "rowSpan", 1)
                if battr(tc, "hMerge") or battr(tc, "vMerge"):
                    c += 1
                    continue
                tcPr = kid(tc, "tcPr")
                fill, extra, bds = None, {}, {}
                if has_style:
                    for reg in self._regions(tblPr, r, c, nr, nc):
                        d = self._style_region(style, reg, part)
                        if "fill" in d:
                            fill = d["fill"]
                        if "color" in d:
                            extra["fill"] = ("solid", d["color"])
                        for k in ("b", "i"):
                            if k in d:
                                extra[k] = d[k]
                        bd = d.get("bd")
                        if bd:
                            rowreg, colreg = reg in _ROW_REGIONS, reg in _COL_REGIONS
                            top = bd.get("top") if (rowreg or r == 0) else bd.get("insideH")
                            bot = bd.get("bottom") if (rowreg or r + rspan >= nr) else bd.get("insideH")
                            lef = bd.get("left") if (colreg or c == 0) else bd.get("insideV")
                            rig = bd.get("right") if (colreg or c + span >= nc) else bd.get("insideV")
                            for k, v in (("top", top), ("bottom", bot), ("left", lef), ("right", rig)):
                                if v is not None:
                                    bds[k] = v
                f = self.fill(tcPr, part)
                if f is not None:
                    fill = f
                for k, tag in (("left", "lnL"), ("right", "lnR"), ("top", "lnT"), ("bottom", "lnB")):
                    ln = kid(tcPr, tag)
                    if ln is not None:
                        bds[k] = self.line_props(ln, part)
                st = ["padding:%s %s %s %s" % (self.u(iattr(tcPr, "marT", 45720)), self.u(iattr(tcPr, "marR", 91440)),
                                               self.u(iattr(tcPr, "marB", 45720)), self.u(iattr(tcPr, "marL", 91440)))]
                css, cls = self.fill_css(fill)
                if css:
                    st.append(css)
                for k, ln in bds.items():
                    b = self._border(ln)
                    st.append("border-%s:%s" % (k, b or "none"))
                anchor = tcPr.get("anchor") if tcPr is not None else None
                if anchor in ("ctr", "b"):
                    st.append("vertical-align:" + ("middle" if anchor == "ctr" else "bottom"))
                txBody = kid(tc, "txBody")
                inner = self.paragraphs(txBody, part, sources, extra or None, kid(txBody, "lstStyle")) if txBody is not None else ""
                out.append('<td%s%s%s style="%s">%s</td>' % (
                    ' class="%s"' % cls if cls else "", ' colspan="%d"' % span if span > 1 else "",
                    ' rowspan="%d"' % rspan if rspan > 1 else "", esc(";".join(st)), inner))
                c += span
            out.append("</tr>")
        out.append("</table>")
        return "".join(out)


# ---------------------------------------------------------------------------- charts

_BAR, _LINE_K, _AREA, _PIE, _SCAT = "bar", "line", "area", "pie", "scatter"
_CHART_KINDS = {"barChart": _BAR, "bar3DChart": _BAR, "lineChart": _LINE_K, "line3DChart": _LINE_K,
                "stockChart": _LINE_K, "radarChart": _LINE_K, "areaChart": _AREA, "area3DChart": _AREA,
                "pieChart": _PIE, "pie3DChart": _PIE, "doughnutChart": _PIE, "ofPieChart": _PIE,
                "scatterChart": _SCAT, "bubbleChart": _SCAT}


def _cache_pts(ref, numeric):
    """c:cat / c:val / c:xVal ... -> list indexed by point idx."""
    if ref is None:
        return []
    cache = None
    for e in ref.iter():
        if _ln(e) in ("numCache", "strCache", "numLit", "strLit"):
            cache = e
            break
    if cache is None:
        return []
    n = iattr(kid(cache, "ptCount"), "val", 0)
    pts = {}
    for pt in kids(cache, "pt"):
        v = kid(pt, "v")
        txt = v.text if v is not None else None
        i = iattr(pt, "idx")
        if numeric:
            try:
                pts[i] = float(txt)
            except (TypeError, ValueError):
                pts[i] = None
        else:
            pts[i] = txt or ""
    n = max(n, max(pts) + 1 if pts else 0)
    return [pts.get(i, None if numeric else "") for i in range(n)]


def _nice(lo, hi, n=5):
    if hi <= lo:
        hi = lo + 1.0
    raw = (hi - lo) / n
    mag = 10 ** math.floor(math.log10(raw))
    step = next(s * mag for s in (1, 2, 2.5, 5, 10) if s * mag >= raw)
    a, b = math.floor(lo / step) * step, math.ceil(hi / step) * step
    ticks = []
    v = a
    while v <= b + step * 1e-6:
        ticks.append(round(v, 10))
        v += step
    return a, b, ticks


def _fmt(v):
    if abs(v) >= 1e6 or (v and abs(v) < 1e-3):
        return "%.3g" % v
    return ("%.3f" % v).rstrip("0").rstrip(".") if v != int(v) else "{:,}".format(int(v))


class _Charts:
    def chart(self, part, rid, w, h):
        try:
            cpart = part.rels[rid].target_part
            root = etree.fromstring(cpart.blob)
        except Exception:
            return ""
        ch = kid(root, "chart")
        plot = kid(ch, "plotArea")
        if plot is None:
            return ""
        accents = [self.scheme("accent%d" % i) for i in range(1, 7)]
        series, pie_el, grouping, horiz = [], None, "clustered", False
        for ct in plot:
            kind = _CHART_KINDS.get(_ln(ct))
            if not kind:
                continue
            if kind == _BAR:
                horiz = kid(ct, "barDir") is not None and kid(ct, "barDir").get("val") == "bar"
                g = kid(ct, "grouping")
                grouping = g.get("val") if g is not None else "clustered"
            elif kind == _AREA and kid(ct, "grouping") is not None and kid(ct, "grouping").get("val") != "standard":
                grouping = kid(ct, "grouping").get("val")
            elif kind == _PIE:
                pie_el = ct
            show_val = battr(path(ct, "dLbls", "showVal"), "val")
            for ser in kids(ct, "ser"):
                i = len(series)
                spPr = kid(ser, "spPr")
                f = self.fill(spPr, cpart)
                col = self.fill_color(f)
                lf = self.fill(kid(spPr, "ln"), cpart)
                if kind in (_LINE_K, _SCAT):
                    col = self.fill_color(lf) or col
                no_line = bool(lf and lf[0] == "none")
                if col is None:
                    col = accents[i % 6] if i < 6 else accents[i % 6].mix(Clr(0, 0, 0), 0.25)
                sv = battr(path(ser, "dLbls", "showVal"), "val") if kid(ser, "dLbls") is not None else show_val
                series.append({
                    "kind": kind, "name": clean(" ".join(_cache_pts(kid(ser, "tx"), False)) or _text_of(kid(ser, "tx")) or "Series %d" % (i + 1)),
                    "cat": _cache_pts(kid(ser, "cat"), False), "val": _cache_pts(kid(ser, "val"), True),
                    "x": _cache_pts(kid(ser, "xVal"), True), "y": _cache_pts(kid(ser, "yVal"), True),
                    "color": col, "show_val": sv, "no_line": no_line,
                    "marker": not (path(ser, "marker", "symbol") is not None and path(ser, "marker", "symbol").get("val") == "none"),
                    "dpt": {iattr(kid(d, "idx"), "val"): self.fill_color(self.fill(kid(d, "spPr"), cpart)) for d in kids(ser, "dPt")}})
        if not series:
            return ""
        W = 1000.0
        H = W * h / max(w, 1.0)
        fs = min(max(152400.0 / (max(w, 1.0) / W), 9.0), 40.0)  # ~12pt on the slide
        txt = self.scheme("tx1").mix(self.scheme("bg1"), 0.3).css()
        grid = self.scheme("tx1").mix(self.scheme("bg1"), 0.85).css()
        font = esc(font_stack(self.theme.font("+mn-lt")))
        o = []
        top, bottom, left, right = fs * 0.8, H - fs * 0.8, fs * 0.8, W - fs * 0.8
        title = ""
        if kid(ch, "title") is not None and not battr(kid(ch, "autoTitleDeleted"), "val"):
            title = clean(_text_of(kid(ch, "title"))) or (series[0]["name"] if len(series) == 1 else "")
        if title:
            o.append('<text x="%s" y="%s" text-anchor="middle" font-size="%s" fill="%s">%s</text>' % (
                num(W / 2), num(top + fs * 1.1), num(fs * 1.4), txt, esc(title)))
            top += fs * 2.4
        is_pie = series[0]["kind"] == _PIE
        legend = kid(ch, "legend")
        if legend is not None:
            if is_pie:
                cats = series[0]["cat"] or [str(i + 1) for i in range(len(series[0]["val"]))]
                items = [(clean(cname), series[0]["dpt"].get(i) or accents[i % 6].mix(Clr(1, 1, 1), 0.35 * (i // 6 % 2)))
                         for i, cname in enumerate(cats)]
            else:
                items = [(s["name"], s["color"]) for s in series]
            pos = path(legend, "legendPos")
            if pos is not None and pos.get("val") in ("r", "l", "tr") or (pos is None):
                lw = min(max(len(t) for t, _c in items) * fs * 0.58 + fs * 2.2, W * 0.35)
                y = (top + bottom) / 2 - len(items) * fs * 0.8
                for t, c in items:
                    o.append('<rect x="%s" y="%s" width="%s" height="%s" fill="%s"/><text x="%s" y="%s" font-size="%s" fill="%s">%s</text>' % (
                        num(right - lw), num(y), num(fs * .8), num(fs * .8), c.css(), num(right - lw + fs * 1.2), num(y + fs * .75), num(fs), txt, esc(t)))
                    y += fs * 1.6
                right -= lw + fs * 0.6
            else:
                widths = [len(t) * fs * 0.58 + fs * 2.4 for t, _c in items]
                x = max((W - sum(widths)) / 2, left)
                y = bottom - fs
                for (t, c), ww in zip(items, widths):
                    o.append('<rect x="%s" y="%s" width="%s" height="%s" fill="%s"/><text x="%s" y="%s" font-size="%s" fill="%s">%s</text>' % (
                        num(x), num(y), num(fs * .8), num(fs * .8), c.css(), num(x + fs * 1.2), num(y + fs * .75), num(fs), txt, esc(t)))
                    x += ww
                bottom -= fs * 2.2
        if is_pie:
            o.append(self._pie(series[0], pie_el, accents, left, top, right, bottom, fs))
        else:
            o.append(self._axes_chart(series, grouping, horiz, left, top, right, bottom, fs, txt, grid))
        return '<svg class="ch" viewBox="0 0 %s %s" font-family="%s">%s</svg>' % (num(W), num(H), font, "".join(o))

    def _pie(self, s, pie_el, accents, left, top, right, bottom, fs):
        vals = [v if v and v > 0 else 0.0 for v in s["val"]]
        total = sum(vals) or 1.0
        cx, cy = (left + right) / 2, (top + bottom) / 2
        R = max(min(right - left, bottom - top) / 2 - fs * 0.5, 1.0)
        hole = 0.0
        if pie_el is not None and _ln(pie_el) == "doughnutChart":
            hole = iattr(kid(pie_el, "holeSize"), "val", 50) / 100.0 * R
        pct = battr(path(pie_el, "dLbls", "showPercent"), "val") or s["show_val"]
        a = -math.pi / 2 + math.radians(iattr(kid(pie_el, "firstSliceAng"), "val", 0))
        o = []
        for i, v in enumerate(vals):
            if v <= 0:
                continue
            b = a + 2 * math.pi * v / total
            col = s["dpt"].get(i) or accents[i % 6].mix(Clr(1, 1, 1), 0.35 * (i // 6 % 2))
            big = 1 if b - a > math.pi else 0
            if v >= total * 0.9999:
                d = "M%s %sA%s %s 0 1 1 %s %sA%s %s 0 1 1 %s %sZ" % tuple(num(x) for x in (cx - R, cy, R, R, cx + R, cy, R, R, cx - R, cy))
            else:
                p = [cx + R * math.cos(a), cy + R * math.sin(a), cx + R * math.cos(b), cy + R * math.sin(b)]
                if hole:
                    q = [cx + hole * math.cos(b), cy + hole * math.sin(b), cx + hole * math.cos(a), cy + hole * math.sin(a)]
                    d = "M%s %sA%s %s 0 %d 1 %s %sL%s %sA%s %s 0 %d 0 %s %sZ" % (
                        num(p[0]), num(p[1]), num(R), num(R), big, num(p[2]), num(p[3]), num(q[0]), num(q[1]), num(hole), num(hole), big, num(q[2]), num(q[3]))
                else:
                    d = "M%s %sL%s %sA%s %s 0 %d 1 %s %sZ" % (num(cx), num(cy), num(p[0]), num(p[1]), num(R), num(R), big, num(p[2]), num(p[3]))
            o.append('<path d="%s" fill="%s" stroke="%s" stroke-width="%s"/>' % (d, col.css(), self.scheme("bg1").css(), num(fs * .12)))
            if pct and v / total > 0.04:
                m, rr = (a + b) / 2, (R + hole) / 2 if hole else R * 0.65
                o.append('<text x="%s" y="%s" text-anchor="middle" font-size="%s" fill="%s">%s</text>' % (
                    num(cx + rr * math.cos(m)), num(cy + rr * math.sin(m) + fs * .35), num(fs),
                    "#fff" if col.lum() < 0.55 else "#000", "%d%%" % round(v / total * 100)))
            a = b
        return "".join(o)

    def _axes_chart(self, series, grouping, horiz, left, top, right, bottom, fs, txt, grid):
        scatter = series[0]["kind"] == _SCAT
        stacked = grouping in ("stacked", "percentStacked")
        pctst = grouping == "percentStacked"
        ncat = max([len(s["val"]) for s in series] + [len(s["cat"]) for s in series] + [1])
        cats = next((s["cat"] for s in series if s["cat"]), [str(i + 1) for i in range(ncat)])
        o = []
        if scatter:
            xs = [v for s in series for v in s["x"] if v is not None] or [0.0, 1.0]
            ys = [v for s in series for v in s["y"] if v is not None] or [0.0, 1.0]
            x0, x1, xt = _nice(min(xs), max(xs))
            lo, hi, ticks = _nice(min(ys + [0.0]) if min(ys) >= 0 else min(ys), max(ys))
        else:
            tot_pos, tot_neg = [0.0] * ncat, [0.0] * ncat
            allv = [0.0]
            for s in series:
                for i, v in enumerate(s["val"][:ncat]):
                    if v is None:
                        continue
                    allv.append(v)
                    if s["kind"] in (_BAR, _AREA):
                        if v >= 0:
                            tot_pos[i] += v
                        else:
                            tot_neg[i] += v
            if pctst:
                lo, hi, ticks = 0.0, 1.0, [0, .2, .4, .6, .8, 1.0]
            elif stacked:
                lo, hi, ticks = _nice(min(tot_neg + [0.0]), max(tot_pos + [0.0]))
            else:
                lo, hi, ticks = _nice(min(allv), max(allv))
        labels = [("%d%%" % round(t * 100)) if (pctst and not scatter) else _fmt(t) for t in ticks]
        if horiz:
            left += max(len(clean(str(c))) for c in cats) * fs * 0.55 + fs * 0.6
            bottom -= fs * 1.6
            right -= fs
        else:
            left += max(len(l) for l in labels) * fs * 0.58 + fs * 0.6
            bottom -= fs * 1.7
            top += fs * 0.5
        pw, ph = max(right - left, 1.0), max(bottom - top, 1.0)

        def vy(v):  # value -> pixel along the value axis
            f = (v - lo) / (hi - lo) if hi > lo else 0.0
            return left + f * pw if horiz else bottom - f * ph

        for t, lab in zip(ticks, labels):
            p = vy(t)
            if horiz:
                o.append('<line x1="%s" y1="%s" x2="%s" y2="%s" stroke="%s"/><text x="%s" y="%s" text-anchor="middle" font-size="%s" fill="%s">%s</text>' % (
                    num(p), num(top), num(p), num(bottom), grid, num(p), num(bottom + fs * 1.2), num(fs), txt, lab))
            else:
                o.append('<line x1="%s" y1="%s" x2="%s" y2="%s" stroke="%s"/><text x="%s" y="%s" text-anchor="end" font-size="%s" fill="%s">%s</text>' % (
                    num(left), num(p), num(right), num(p), grid, num(left - fs * .4), num(p + fs * .35), num(fs), txt, lab))
        if scatter:
            for t in xt:
                x = left + (t - x0) / (x1 - x0) * pw
                o.append('<text x="%s" y="%s" text-anchor="middle" font-size="%s" fill="%s">%s</text>' % (num(x), num(bottom + fs * 1.2), num(fs), txt, _fmt(t)))
            for s in series:
                pts = [(left + (x - x0) / (x1 - x0) * pw, vy(y)) for x, y in zip(s["x"], s["y"]) if x is not None and y is not None]
                if not s["x"]:
                    pts = [(left + (i + .5) / max(len(s["y"]), 1) * pw, vy(y)) for i, y in enumerate(s["y"]) if y is not None]
                c = s["color"].css()
                if len(pts) > 1 and not s["no_line"]:
                    o.append('<polyline points="%s" fill="none" stroke="%s" stroke-width="%s" stroke-linejoin="round"/>' % (
                        " ".join("%s,%s" % (num(x, 1), num(y, 1)) for x, y in pts), c, num(fs * .18)))
                if s["marker"] and len(pts) <= 400:
                    o += ['<circle cx="%s" cy="%s" r="%s" fill="%s"/>' % (num(x, 1), num(y, 1), num(fs * .28), c) for x, y in pts]
            return "".join(o)
        band = (ph if horiz else pw) / ncat
        step = max(1, int(math.ceil(ncat * fs * (1.2 if horiz else 3.0) / (ph if horiz else pw)))) if ncat > 1 else 1
        for i in range(0, ncat, step):
            lab = esc(clean(str(cats[i] if i < len(cats) else i + 1))[:28])
            if horiz:
                y = bottom - (i + .5) * band
                o.append('<text x="%s" y="%s" text-anchor="end" font-size="%s" fill="%s">%s</text>' % (num(left - fs * .4), num(y + fs * .35), num(fs), txt, lab))
            else:
                o.append('<text x="%s" y="%s" text-anchor="middle" font-size="%s" fill="%s">%s</text>' % (num(left + (i + .5) * band), num(bottom + fs * 1.3), num(fs), txt, lab))
        bars = [s for s in series if s["kind"] == _BAR]
        acc_p, acc_n = [0.0] * ncat, [0.0] * ncat
        sums = [sum(abs(s["val"][i]) for s in bars if i < len(s["val"]) and s["val"][i] is not None) or 1.0 for i in range(ncat)]
        nb = 1 if stacked else max(len(bars), 1)
        bw = band * 0.68 / nb
        for k, s in enumerate(bars):
            c = s["color"].css()
            for i, v in enumerate(s["val"][:ncat]):
                if v is None:
                    continue
                vv = v / sums[i] if pctst else v
                if stacked:
                    base = acc_p[i] if vv >= 0 else acc_n[i]
                    if vv >= 0:
                        acc_p[i] += vv
                    else:
                        acc_n[i] += vv
                    a, b = vy(base), vy(base + vv)
                    offp = band * 0.16
                else:
                    a, b = vy(max(lo, 0.0) if lo <= 0 <= hi else lo), vy(vv)
                    offp = band * 0.16 + k * bw
                col = (s["dpt"].get(i).css() if s["dpt"].get(i) is not None else c)
                if horiz:
                    y = bottom - i * band - offp - bw
                    o.append('<rect x="%s" y="%s" width="%s" height="%s" fill="%s"/>' % (num(min(a, b), 1), num(y, 1), num(abs(b - a), 1), num(bw, 1), col))
                    if s["show_val"]:
                        o.append('<text x="%s" y="%s" font-size="%s" fill="%s">%s</text>' % (num(max(a, b) + fs * .3), num(y + bw / 2 + fs * .35), num(fs * .9), txt, _fmt(v)))
                else:
                    x = left + i * band + offp
                    o.append('<rect x="%s" y="%s" width="%s" height="%s" fill="%s"/>' % (num(x, 1), num(min(a, b), 1), num(bw, 1), num(abs(b - a), 1), col))
                    if s["show_val"]:
                        o.append('<text x="%s" y="%s" text-anchor="middle" font-size="%s" fill="%s">%s</text>' % (num(x + bw / 2), num(min(a, b) - fs * .3), num(fs * .9), txt, _fmt(v)))
        acc = [0.0] * ncat
        for s in series:
            if s["kind"] not in (_LINE_K, _AREA):
                continue
            c = s["color"].css()
            pts = []
            for i, v in enumerate(s["val"][:ncat]):
                if v is None:
                    continue
                if s["kind"] == _AREA and stacked:
                    acc[i] += v / (sums[i] if pctst and bars else 1.0)
                    v = acc[i]
                pos = (i + .5) * band
                pts.append((vy(v), bottom - pos) if horiz else (left + pos, vy(v)))
            if not pts:
                continue
            pl = " ".join("%s,%s" % (num(x, 1), num(y, 1)) for x, y in pts)
            if s["kind"] == _AREA:
                z = vy(max(lo, 0.0))
                o.append('<polygon points="%s,%s %s %s,%s" fill="%s" fill-opacity=".8"/>' % (num(pts[0][0], 1), num(z, 1), pl, num(pts[-1][0], 1), num(z, 1), c))
            else:
                if not s["no_line"]:
                    o.append('<polyline points="%s" fill="none" stroke="%s" stroke-width="%s" stroke-linejoin="round" stroke-linecap="round"/>' % (pl, c, num(fs * .2)))
                if s["marker"] and len(pts) <= 60:
                    o += ['<circle cx="%s" cy="%s" r="%s" fill="%s"/>' % (num(x, 1), num(y, 1), num(fs * .3), c) for x, y in pts]
        o.append('<line x1="%s" y1="%s" x2="%s" y2="%s" stroke="%s" stroke-width="1.5"/>' % (
            (num(left), num(top), num(left), num(bottom), txt) if horiz else (num(left), num(bottom), num(right), num(bottom), txt)))
        return "".join(o)


# ------------------------------------------------------------------------------ deck

_CSS = """
.{u}{{display:block;margin:0;padding:0}}
.{u} .sf{{container-type:inline-size;max-width:{maxw};margin:0 auto 14px;background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.25);border-radius:2px;overflow:hidden}}
.{u} .sl{{position:relative;width:100cqw;height:{ratio}cqw;overflow:hidden;background:#fff;background-size:100% 100%;line-height:1.2;text-align:left;color:#000;-webkit-text-size-adjust:100%;text-size-adjust:100%;font-kerning:normal}}
.{u} .s{{position:absolute}}
.{u} .g{{position:absolute;overflow:visible;pointer-events:none}}
.{u} .im{{position:absolute;inset:0;background-size:100% 100%;background-repeat:no-repeat}}
.{u} .missing{{background:#eee;outline:1px dashed #bbb}}
.{u} .t{{position:absolute;display:flex;flex-direction:column;box-sizing:border-box;white-space:pre-wrap;overflow-wrap:break-word;tab-size:4}}
.{u} .nw{{white-space:pre}}
.{u} p{{margin:0;padding:0;line-height:1.2;min-width:0;border:0}}
.{u} .tab{{white-space:pre}}
.{u} a{{color:inherit}}
.{u} .tb{{position:absolute;left:0;top:0;width:100%;border-collapse:collapse;table-layout:fixed;border-spacing:0}}
.{u} .tb td{{vertical-align:top;white-space:pre-wrap;overflow-wrap:break-word;box-sizing:border-box;border:0;background-size:100% 100%}}
.{u} .ch{{position:absolute;inset:0;width:100%;height:100%}}
.{u} .nt{{max-width:{maxw};margin:-8px auto 14px;font:13px/1.45 system-ui,-apple-system,'Segoe UI',sans-serif;color:#555}}
.{u} .nt summary{{cursor:pointer;color:#777}}
.{u} .nt div{{white-space:pre-wrap;padding:4px 0 0 14px}}
@media (prefers-color-scheme:dark){{.{u} .nt,.{u} .nt summary{{color:#aaa}}}}
.{u} .bar{{display:none}}
.{u}.dk .pg{{display:none}}
.{u}.dk .pg.on{{display:block}}
.{u}.dk .sf{{margin-bottom:8px}}
.{u}.dk .nt{{margin-top:0}}
.{u}.dk .bar{{display:flex;align-items:center;gap:2px;width:max-content;max-width:100%;margin:0 auto 10px;padding:3px;box-sizing:border-box;font:500 13px/1 Inter,ui-sans-serif,system-ui,-apple-system,'Segoe UI',sans-serif;color:#4d4d4d;background:#fff;border:1px solid #e3e3e3;border-radius:12px;box-shadow:0 1px 2px rgba(0,0,0,.04);user-select:none;-webkit-user-select:none}}
.{u} .bar button{{display:inline-flex;align-items:center;justify-content:center;gap:6px;height:28px;min-width:28px;padding:0 8px;font:inherit;color:inherit;background:transparent;border:0;border-radius:8px;cursor:pointer;transition:background .15s ease,color .15s ease}}
.{u} .bar button:hover{{background:#ececec;color:#171717}}
.{u} .bar button:active{{background:#e3e3e3}}
.{u} .bar button:focus-visible{{outline:2px solid rgba(120,120,120,.55);outline-offset:1px}}
.{u} .bar button:disabled{{opacity:.3;cursor:default;background:transparent}}
.{u} .bar svg{{width:15px;height:15px;flex:none;fill:none;stroke:currentColor;stroke-width:2;stroke-linecap:round;stroke-linejoin:round}}
.{u} .bar .n{{min-width:4.2em;padding:0 4px;text-align:center;font-variant-numeric:tabular-nums;color:#676767;letter-spacing:.01em}}
.{u} .bar .sp{{width:1px;height:16px;background:#e3e3e3;margin:0 4px}}
@media (prefers-color-scheme:dark){{.{u}.dk .bar{{color:#cdcdcd;background:#262626;border-color:#333;box-shadow:0 1px 2px rgba(0,0,0,.3)}}
.{u} .bar button:hover{{background:#333;color:#fff}}.{u} .bar button:active{{background:#3d3d3d}}
.{u} .bar .n{{color:#9b9b9b}}.{u} .bar .sp{{background:#3d3d3d}}}}
@media print{{.{u} .sf{{box-shadow:none;break-inside:avoid;margin-bottom:0}}}}
"""


class _SlideCtx(_Paint, _Text, _Shapes, _Tables, _Charts):
    def __init__(self, deck, slide, no):
        self.deck = deck
        self.slide_no = no
        self.part = slide.part
        self.slide_el = slide._element
        layout = slide.slide_layout
        master = layout.slide_master
        self.layout_el, self.layout_part = layout._element, layout.part
        self.master_el, self.master_part = master._element, master.part
        self.theme = deck.theme_for(master.part)
        cm = dict(_DEFAULT_CLRMAP)
        m = kid(self.master_el, "clrMap")
        if m is not None:
            cm.update(m.attrib)
        for el in (self.layout_el, self.slide_el):
            ov = path(el, "clrMapOvr", "overrideClrMapping")
            if ov is not None:
                cm.update(ov.attrib)
        self.clrmap = deck.intern_clrmap(cm)

    def u(self, emu):
        return num(emu / self.deck.sw * 100.0, 3) + "cqw"

    def background(self):
        for el, part in ((self.slide_el, self.part), (self.layout_el, self.layout_part), (self.master_el, self.master_part)):
            bg = path(el, "cSld", "bg")
            if bg is None:
                continue
            f = self.fill(kid(bg, "bgPr"), part)
            if f is None:
                ref = kid(bg, "bgRef")
                if ref is not None:
                    c = self.color(color_el(ref))
                    idx = iattr(ref, "idx")
                    lst = self.theme.bgfills if idx >= 1001 else self.theme.fills
                    i = idx - 1001 if idx >= 1001 else idx - 1
                    if 0 <= i < len(lst):
                        holder = etree.Element("x")
                        holder.append(copy.deepcopy(lst[i]))
                        f = self.fill(holder, self.theme.part or self.master_part, c)
                        if f == NONE_FILL:
                            f = None
                    if f is None and c is not None:
                        f = ("solid", c)
            if f is not None:
                return f
        return ("solid", Clr(1, 1, 1))

    def render(self):
        css, cls = self.fill_css(self.background(), slide_bg=True)
        body = []
        if self.slide_el.get("showMasterSp") not in ("0", "false"):
            if self.layout_el.get("showMasterSp") not in ("0", "false"):
                body.append(self.tree(path(self.master_el, "cSld", "spTree"), self.master_part, IDENT, "master"))
            body.append(self.tree(path(self.layout_el, "cSld", "spTree"), self.layout_part, IDENT, "layout"))
        body.append(self.tree(path(self.slide_el, "cSld", "spTree"), self.part, IDENT, "slide"))
        return '<div class="sl%s" style="%s" role="img" aria-label="Slide %d">%s</div>' % (
            " " + cls if cls else "", esc(css), self.slide_no, "".join(body))


class _Deck:
    def __init__(self, src, max_image_px=2400, recompress_over=400_000):
        if isinstance(src, (bytes, bytearray)):
            data = bytes(src)
        elif hasattr(src, "read"):
            data = src.read()
        else:
            with open(src, "rb") as fh:
                data = fh.read()
        self.uid = "px%08x" % (zlib.crc32(data) & 0xFFFFFFFF)
        self.prs = Presentation(io.BytesIO(data))
        self.sw = float(self.prs.slide_width or DEFAULT_W)
        self.sh = float(self.prs.slide_height or DEFAULT_H)
        self.images = _Images(self, max_image_px, recompress_over)
        self.lvl_cache, self.ph_cache, self.tbl_cache = {}, {}, {}
        self._themes, self._clrmaps, self._ids = {}, {}, 0
        self.warnings = []
        self.default_text = kid(self.prs.part._element, "defaultTextStyle")
        self.table_styles = {}
        try:
            ts = self.prs.part.part_related_by(RT.TABLE_STYLES)
            for st in etree.fromstring(ts.blob):
                if st.get("styleId"):
                    self.table_styles[st.get("styleId")] = st
        except Exception:
            pass

    def warn(self, msg):
        self.warnings.append(msg)

    def new_id(self, prefix):
        self._ids += 1
        return "%s%s%d" % (self.uid, prefix, self._ids)

    def intern_clrmap(self, cm):
        key = tuple(sorted(cm.items()))
        return self._clrmaps.setdefault(key, cm)

    def theme_for(self, master_part):
        key = str(master_part.partname)
        if key not in self._themes:
            try:
                tp = master_part.part_related_by(RT.THEME)
                self._themes[key] = Theme(etree.fromstring(tp.blob), tp)
            except Exception:
                self._themes[key] = Theme()
        return self._themes[key]

    def slides_html(self, notes=True):
        out = []
        for no, slide in enumerate(self.prs.slides, 1):
            try:
                body = _SlideCtx(self, slide, no).render()
            except Exception as e:
                self.warn("slide %d failed: %s: %s" % (no, type(e).__name__, e))
                body = '<div class="sl" style="display:flex;align-items:center;justify-content:center;font:3cqw sans-serif;color:#888">Slide %d could not be rendered</div>' % no
            note = ""
            if notes:
                try:
                    if slide.has_notes_slide:
                        tf = slide.notes_slide.notes_text_frame
                        txt = clean(tf.text).strip() if tf is not None else ""
                        if txt:
                            note = '<details class="nt"><summary>Notes &middot; slide %d</summary><div>%s</div></details>' % (no, esc(txt))
                except Exception:
                    pass
            out.append((body, note))
        return out


def render_slides(src, notes=False, max_image_px=2400):
    """-> (css, [slide_html, ...], warnings). Each slide_html needs the css and a wrapper with the deck class."""
    deck = _Deck(src, max_image_px)
    slides = deck.slides_html(notes)
    return _deck_css(deck, "1280px"), ['<div class="%s"><div class="sf">%s</div>%s</div>' % (deck.uid, b, n) for b, n in slides], deck.warnings


def _deck_css(deck, max_width):
    css = _CSS.format(u=deck.uid, ratio=num(deck.sh / deck.sw * 100.0, 4), maxw=max_width)
    return css.strip().replace("\n", "") + "".join(deck.images.css)


_DECK_JS = """(function(){var r=document.currentScript.parentElement,p=r.querySelectorAll('.pg'),n=p.length,i=0,all=false;
var b=r.querySelector('.bar'),prev=b.querySelector('.pv'),next=b.querySelector('.nx'),lab=b.querySelector('.n'),tg=b.querySelector('.al .tx'),tb=b.querySelector('.al'),sp=b.querySelector('.sp');
function h(){try{parent.postMessage({type:'iframe:height',height:document.documentElement.scrollHeight},'*')}catch(e){}}
function show(k){i=Math.max(0,Math.min(n-1,k));for(var j=0;j<n;j++)p[j].classList.toggle('on',j===i);
lab.textContent=(i+1)+' / '+n;prev.disabled=i===0;next.disabled=i===n-1;h()}
function mode(a){all=a;r.classList.toggle('dk',!all);tg.textContent=all?'One at a time':'Show all';
prev.style.display=next.style.display=lab.style.display=sp.style.display=all?'none':'';h()}
prev.onclick=function(){show(i-1)};next.onclick=function(){show(i+1)};tb.onclick=function(){mode(!all)};
document.addEventListener('keydown',function(e){if(all)return;if(e.key==='ArrowRight'||e.key==='PageDown'){show(i+1);e.preventDefault()}
else if(e.key==='ArrowLeft'||e.key==='PageUp'){show(i-1);e.preventDefault()}else if(e.key==='Home')show(0);else if(e.key==='End')show(n-1)});
mode(false);show(0);window.addEventListener('load',h);window.addEventListener('resize',h);
if(window.ResizeObserver)new ResizeObserver(h).observe(document.body)})();"""
_HEIGHT_JS = ("(function(){function h(){try{parent.postMessage({type:'iframe:height',height:document.documentElement.scrollHeight},'*')}"
              "catch(e){}}window.addEventListener('load',h);if(window.ResizeObserver)new ResizeObserver(h).observe(document.body)})();")


def convert(src, full_document=True, notes=True, title=None, max_width="1280px", max_image_px=2400, warnings=None,
            viewer="scroll"):
    """Convert a .pptx (path, bytes or binary file object) to an HTML string containing every slide.

    viewer="scroll" stacks all slides; viewer="deck" shows one slide at a time with Previous / Next
    buttons, arrow-key navigation and a "Show all" toggle (falls back to the stacked view without JS).
    full_document=False returns an embeddable <div> fragment (with its own scoped <style>).
    Both viewers report their height to a parent frame (Open WebUI embeds) via postMessage.
    Pass a list as `warnings` to collect non-fatal rendering notes.
    """
    deck = _Deck(src, max_image_px)
    slides = deck.slides_html(notes)
    body = "".join('<div class="pg"><div class="sf">%s</div>%s</div>' % (b, n) for b, n in slides)
    if warnings is not None:
        warnings.extend(deck.warnings)
    bar = ""
    script = _HEIGHT_JS
    if viewer == "deck" and len(slides) > 1:
        ico = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="%s"/></svg>'
        bar = ('<div class="bar"><button class="pv" type="button" title="Previous slide" aria-label="Previous slide">%s</button>'
               '<span class="n"></span><button class="nx" type="button" title="Next slide" aria-label="Next slide">%s</button>'
               '<span class="sp"></span><button class="al" type="button"><span class="tx">Show all</span></button></div>') % (
                   ico % "M15 6l-6 6 6 6", ico % "M9 6l6 6-6 6")
        script = _DECK_JS
    frag = '<div class="%s"><style>%s</style>%s%s<script>%s</script></div>' % (
        deck.uid, _deck_css(deck, max_width), bar, body, script)
    if not full_document:
        return frag
    if title is None:
        title = deck.prs.core_properties.title or (src if isinstance(src, str) else "Presentation")
        title = str(title).rsplit("/", 1)[-1]
    page = ("html,body{margin:0}body{padding:8px;background:transparent}" if viewer == "deck" else
            "html,body{margin:0}body{background:#f1f1f1;padding:14px}@media (prefers-color-scheme:dark){body{background:#1e1e1e}}")
    return ('<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width,initial-scale=1"><title>%s</title>'
            '<style>%s</style></head><body>%s</body></html>') % (esc(clean(title)), page, frag)


def slide_text(src):
    """Plain text of every slide (for search / LLM context): [{"slide": n, "text": ..., "notes": ...}]."""
    if isinstance(src, (bytes, bytearray)):
        src = io.BytesIO(bytes(src))
    out = []
    for no, slide in enumerate(Presentation(src).slides, 1):
        lines = []
        for t in slide._element.iter(qa("p")):
            s = clean("".join(x.text or "" for x in t.iter(qa("t")))).strip()
            if s:
                lines.append(s)
        note = ""
        try:
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame is not None:
                note = clean(slide.notes_slide.notes_text_frame.text).strip()
        except Exception:
            pass
        out.append({"slide": no, "text": "\n".join(lines), "notes": note})
    return out


def convert_file(src_path, out_path=None, **kw):
    out_path = out_path or re.sub(r"\.pptx?$", "", src_path, flags=re.I) + ".html"
    html = convert(src_path, **kw)
    with open(out_path, "w", encoding="utf-8") as fh:
        fh.write(html)
    return out_path


def main(argv=None):
    import argparse
    ap = argparse.ArgumentParser(description="Convert .pptx to self-contained HTML (all slides).")
    ap.add_argument("pptx")
    ap.add_argument("-o", "--output")
    ap.add_argument("--fragment", action="store_true", help="emit an embeddable <div> instead of a full document")
    ap.add_argument("--no-notes", action="store_true")
    ap.add_argument("--deck", action="store_true", help="one slide at a time with Previous / Next controls")
    ap.add_argument("--max-image-px", type=int, default=2400)
    a = ap.parse_args(argv)
    warns = []
    t = time.time()
    out = convert_file(a.pptx, a.output, full_document=not a.fragment, notes=not a.no_notes,
                       max_image_px=a.max_image_px, warnings=warns,
                       viewer="deck" if a.deck else "scroll")
    for w in warns:
        sys.stderr.write("warning: %s\n" % w)
    sys.stderr.write("%s  (%.2fs)\n" % (out, time.time() - t))



# ECMA-376 preset shape geometry (zlib + base64 JSON), see tools/build_presets.py
_PRESET_B64 = """
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
/f9zcg7v
"""

if __name__ == "__main__":
    main()
