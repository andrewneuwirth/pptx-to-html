

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
