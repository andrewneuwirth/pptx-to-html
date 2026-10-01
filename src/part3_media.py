

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
