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
