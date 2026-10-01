

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
