

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
