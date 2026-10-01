

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
