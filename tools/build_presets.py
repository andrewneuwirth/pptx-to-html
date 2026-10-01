"""Compile ECMA-376 presetShapeDefinitions.xml into the compact blob embedded in pptx_to_html.py.

Usage: python tools/build_presets.py presetShapeDefinitions.xml > presets.b64
"""
import base64, json, sys, zlib
from lxml import etree

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def gds(parent):
    return [[g.get("name"), g.get("fmla")] for g in parent.iter(A + "gd")] if parent is not None else []


def pt(el):
    p = el.find(A + "pt")
    return [p.get("x"), p.get("y")]


out = {}
root = etree.parse(sys.argv[1]).getroot()
for shp in root:
    if not isinstance(shp.tag, str):
        continue
    d = {"av": gds(shp.find(A + "avLst")), "gd": gds(shp.find(A + "gdLst"))}
    r = shp.find(A + "rect")
    if r is not None:
        d["rect"] = [r.get("l"), r.get("t"), r.get("r"), r.get("b")]
    paths = []
    for p in shp.find(A + "pathLst"):
        if not isinstance(p.tag, str):
            continue
        cmds = []
        for c in p:
            if not isinstance(c.tag, str):
                continue
            n = etree.QName(c).localname
            if n == "moveTo":
                cmds.append(["M"] + pt(c))
            elif n == "lnTo":
                cmds.append(["L"] + pt(c))
            elif n == "arcTo":
                cmds.append(["A", c.get("wR"), c.get("hR"), c.get("stAng"), c.get("swAng")])
            elif n == "cubicBezTo":
                cmds.append(["C"] + sum(([q.get("x"), q.get("y")] for q in c.findall(A + "pt")), []))
            elif n == "quadBezTo":
                cmds.append(["Q"] + sum(([q.get("x"), q.get("y")] for q in c.findall(A + "pt")), []))
            elif n == "close":
                cmds.append(["Z"])
        pd = {"c": cmds}
        for k in ("w", "h", "fill", "stroke", "extrusionOk"):
            if p.get(k) is not None and k != "extrusionOk":
                pd[k] = p.get(k)
        paths.append(pd)
    d["p"] = paths
    out[etree.QName(shp).localname] = d

blob = zlib.compress(json.dumps(out, separators=(",", ":")).encode(), 9)
b64 = base64.b64encode(blob).decode()
sys.stderr.write("%d shapes, %d bytes b64\n" % (len(out), len(b64)))
print("\n".join(b64[i:i + 100] for i in range(0, len(b64), 100)))
