

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


