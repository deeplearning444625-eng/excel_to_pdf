"""
Font replacement done directly inside the .xlsx (zip + XML), without openpyxl.

 * styles.xml   : font "FROM" is renamed to "TO" (every cell style using it follows)
 * strings      : cells whose style uses FROM and that contain Khmer letters get a
                  rich-text copy: Khmer runs -> Khmer font, the rest -> TO
                  (copies are appended, so other cells sharing the text are untouched)
 * drawings / comments / charts : font name replaced in place
Everything else (pictures, shapes, charts, formulas, print setup) is copied byte for byte.
"""
import html
import re
import zipfile

KH = r"[\u1780-\u17FF\u19E0-\u19FF]"
HAS_KHMER = re.compile(KH)
KHMER_RUN = re.compile(rf"{KH}+(?:[ \u200b\u200c\u200d]+{KH}+)*")

FONTS_BLOCK = re.compile(r"<fonts\b[^>]*>(.*?)</fonts>", re.S)
FONT_RE = re.compile(r"<font(?:\s[^>]*)?(?:/>|>.*?</font>)", re.S)
XFS_BLOCK = re.compile(r"<cellXfs\b[^>]*>(.*?)</cellXfs>", re.S)
XF_RE = re.compile(r"<xf\b[^>]*?(?:/>|>.*?</xf>)", re.S)
SI_RE = re.compile(r"<si>.*?</si>|<si/>", re.S)
CELL_RE = re.compile(r"<c\b([^>]*?)(?:/>|>(.*?)</c>)", re.S)
RUN_RE = re.compile(r"<r>\s*(<rPr>.*?</rPr>)?\s*<t\b[^>]*>(.*?)</t>\s*</r>", re.S)
T_RE = re.compile(r"<t\b[^>]*>(.*?)</t>", re.S)
RFONT_RE = re.compile(r'<rFont\s+val="([^"]*)"')
KHMER_ROW_FACTOR = 2.2   # Khmer fonts are tall: row height = font size x this, so tops are not cut
ROW_RE = re.compile(r"<row\b([^>]*?)(/?)>")
ROWNUM_RE = re.compile(r'\sr="[A-Z]+(\d+)"')
S_ATTR = re.compile(r'\ss="(\d+)"')
T_ATTR = re.compile(r'\st="(\w+)"')


def _esc_attr(s: str) -> str:
    return html.escape(s, quote=True)


def _norm(s: str) -> str:
    return html.unescape(s).strip().lower()


# ------------------------------------------------------------ run properties
def _strip_rpr(rpr: str) -> str:
    return re.sub(r"<(family|charset|scheme)\b[^>]*/>", "", rpr)


def rpr_from_font(font_xml: str, name: str) -> str:
    inner = re.sub(r"^<font\b[^>]*?/?>", "", font_xml)
    inner = re.sub(r"</font>$", "", inner)
    inner = re.sub(r"<name\b[^>]*/>", "", inner)
    return _strip_rpr(f'<rPr><rFont val="{_esc_attr(name)}"/>{inner}</rPr>')


def set_font(rpr: str, name: str) -> str:
    rpr = _strip_rpr(rpr)
    if RFONT_RE.search(rpr):
        return RFONT_RE.sub(lambda m: f'<rFont val="{_esc_attr(name)}"', rpr, count=1)
    return rpr.replace("<rPr>", f'<rPr><rFont val="{_esc_attr(name)}"/>', 1)


# ------------------------------------------------------------ rich strings
def _runs(si_inner: str):
    si = re.sub(r"<rPh\b.*?</rPh>|<phoneticPr\b[^>]*/>", "", si_inner, flags=re.S)
    if re.search(r"<r>", si):
        return [(m.group(1), m.group(2), m.group(0)) for m in RUN_RE.finditer(si)]
    m = T_RE.search(si)
    return [(None, m.group(1) if m else "", None)]


def _split(text_raw: str):
    text = html.unescape(text_raw)
    pos = 0
    for m in KHMER_RUN.finditer(text):
        if m.start() > pos:
            yield text[pos:m.start()], False
        yield m.group(), True
        pos = m.end()
    if pos < len(text):
        yield text[pos:], False


def build_si(si_inner, font_xml, from_l, to_font, khmer_font):
    """New inner XML for a string, or None if nothing has to change.
    font_xml = font of the cell using the string (None = no cell context)."""
    runs = _runs(si_inner)

    def rf_of(rpr):
        m = RFONT_RE.search(rpr or "")
        return _norm(m.group(1)) if m else None

    need = False
    for rpr, text, _ in runs:
        rf = rf_of(rpr)
        mine = rf == from_l or (rf is None and font_xml is not None)
        if rf == from_l or (mine and HAS_KHMER.search(html.unescape(text))):
            need = True
    if not need:
        return None

    out = []
    for rpr, text, orig in runs:
        rf = rf_of(rpr)
        if rf == from_l:
            base = set_font(rpr, to_font)
        elif rf is None and font_xml is not None:
            base = set_font(rpr, to_font) if rpr else rpr_from_font(font_xml, to_font)
        else:
            out.append(orig)
            continue
        for piece, is_kh in _split(text):
            r = set_font(base, khmer_font) if is_kh else base
            out.append(f'<r>{r}<t xml:space="preserve">{html.escape(piece, quote=False)}</t></r>')
    return "".join(out)


# ------------------------------------------------------------ styles
def _process_styles(styles: str, from_l: str, to_font: str):
    """-> (new_styles, {fontId: font_xml}, {xfIndex: fontId} for target fonts)"""
    target = {}
    m = FONTS_BLOCK.search(styles)
    if not m:
        return styles, target, {}
    block = m.group(1)
    pieces, last = [], 0
    for i, fm in enumerate(FONT_RE.finditer(block)):
        fx = fm.group(0)
        nm = re.search(r'<name\s+val="([^"]*)"', fx)
        if nm and _norm(nm.group(1)) == from_l:
            target[i] = fx
            fx = fx[:nm.start(1)] + _esc_attr(to_font) + fx[nm.end(1):]
        pieces.append(block[last:fm.start()])
        pieces.append(fx)
        last = fm.end()
    pieces.append(block[last:])
    new_styles = styles[:m.start(1)] + "".join(pieces) + styles[m.end(1):]

    xf_target = {}
    xm = XFS_BLOCK.search(new_styles)
    if xm and target:
        for i, x in enumerate(XF_RE.finditer(xm.group(1))):
            fid = re.search(r'\bfontId="(\d+)"', x.group(0))
            if fid and int(fid.group(1)) in target:
                xf_target[i] = int(fid.group(1))
    return new_styles, target, xf_target


def _bump_rows(text: str, need: dict) -> str:
    """Give rows that hold Khmer text enough height (adds space above the letters)."""
    def cb(m):
        attrs = m.group(1)
        rm = re.search(r'\sr="(\d+)"', attrs)
        want = need.get(int(rm.group(1))) if rm else None
        hm = re.search(r'\sht="([\d.]+)"', attrs)
        if not want or not hm or float(hm.group(1)) >= want:
            return m.group(0)  # no stored height = Excel sizes the row itself
        attrs = attrs[:hm.start()] + f' ht="{want:.2f}"' + attrs[hm.end():]
        attrs = re.sub(r'\scustomHeight="\d"', "", attrs) + ' customHeight="1"'
        return f"<row{attrs}{m.group(2)}>"
    return ROW_RE.sub(cb, text)


# ------------------------------------------------------------ main entry
def needs_font_change(src, from_font: str) -> bool:
    key = re.compile(re.escape(from_font.encode("utf-8")), re.I)  # scans bytes, no lower() copy
    try:
        with zipfile.ZipFile(src) as z:
            for name in z.namelist():
                if name in ("xl/styles.xml", "xl/sharedStrings.xml") or \
                        name.startswith(("xl/drawings/", "xl/comments", "xl/charts/")):
                    if key.search(z.read(name)):
                        return True
    except zipfile.BadZipFile:
        raise
    return False


def replace_fonts_xml(src, dst, from_font, to_font, khmer_font) -> dict:
    from_l = from_font.strip().lower()
    stats = {"sheets": 0, "cells": 0, "runs": 0}
    changed = {}

    with zipfile.ZipFile(src) as zin:
        names = zin.namelist()
        styles = zin.read("xl/styles.xml").decode("utf-8")
        new_styles, target, xf_target = _process_styles(styles, from_l, to_font)
        if new_styles != styles:
            changed["xl/styles.xml"] = new_styles.encode("utf-8")

        # shared strings
        sst = zin.read("xl/sharedStrings.xml").decode("utf-8") if "xl/sharedStrings.xml" in names else None
        si_orig, si_inner, header, footer = [], [], "", ""
        if sst:
            ms = list(SI_RE.finditer(sst))
            if ms:
                header, footer = sst[:ms[0].start()], sst[ms[-1].end():]
                si_orig = [m.group(0) for m in ms]
                si_inner = [s[4:-5] if s.startswith("<si>") else "" for s in si_orig]
            else:
                header, footer = sst, ""
        new_si, cache = [], {}

        def new_string(idx, fid):
            key = (idx, fid)
            if key not in cache:
                inner = build_si(si_inner[idx], target[fid], from_l, to_font, khmer_font) \
                    if idx < len(si_inner) else None
                if inner is None:
                    cache[key] = None
                else:
                    new_si.append(f"<si>{inner}</si>")
                    cache[key] = len(si_orig) + len(new_si) - 1
                    stats["runs"] += 1
            return cache[key]

        # sheets: only touched when they use a target style or name the font in inline text
        pat = re.compile(r'\ss="(?:%s)"' % "|".join(str(i) for i in xf_target)) if xf_target else None
        name_pat = re.compile(re.escape(from_font), re.I)
        for name in names:
            if not (name.startswith("xl/worksheets/") and name.endswith(".xml")):
                continue
            text = zin.read(name).decode("utf-8")
            if not ((pat and pat.search(text)) or name_pat.search(text)):
                continue
            stats["sheets"] += 1
            rows_need = {}

            def note_khmer(attrs, fid):
                rm = ROWNUM_RE.search(" " + attrs)
                if rm:
                    sz = re.search(r'<sz\s+val="([\d.]+)"', target[fid]) if fid is not None else None
                    h = min(409.0, float(sz.group(1) if sz else 11) * KHMER_ROW_FACTOR)
                    rows_need[int(rm.group(1))] = max(rows_need.get(int(rm.group(1)), 0), h)

            def cell_cb(m):
                attrs, body = m.group(1), m.group(2)
                if body is None:
                    return m.group(0)
                if "inlineStr" not in attrs:  # fast path: most cells are not in a target style
                    sm0 = S_ATTR.search(" " + attrs)
                    if not sm0 or int(sm0.group(1)) not in xf_target:
                        return m.group(0)
                sm = S_ATTR.search(" " + attrs)
                fid = xf_target.get(int(sm.group(1))) if sm else None
                tm = T_ATTR.search(" " + attrs)
                t = tm.group(1) if tm else ""
                if fid is not None:
                    stats["cells"] += 1
                if t == "s" and fid is not None:
                    vm = re.search(r"<v>(\d+)</v>", body)
                    if vm:
                        ni = new_string(int(vm.group(1)), fid)
                        if ni is not None:
                            if HAS_KHMER.search(html.unescape(si_inner[int(vm.group(1))])):
                                note_khmer(attrs, fid)
                            nb = body[:vm.start(1)] + str(ni) + body[vm.end(1):]
                            return m.group(0).replace(body, nb, 1)
                elif t == "inlineStr":
                    im = re.search(r"<is>(.*?)</is>", body, re.S)
                    if im:
                        inner = build_si(im.group(1), target[fid] if fid is not None else None,
                                         from_l, to_font, khmer_font)
                        if inner is not None:
                            if HAS_KHMER.search(html.unescape(im.group(1))):
                                note_khmer(attrs, fid)
                            stats["runs"] += 1
                            nb = body[:im.start(1)] + inner + body[im.end(1):]
                            return m.group(0).replace(body, nb, 1)
                return m.group(0)

            new_text = CELL_RE.sub(cell_cb, text)
            if rows_need:
                new_text = _bump_rows(new_text, rows_need)
                stats["rows"] = stats.get("rows", 0) + len(rows_need)
            if new_text != text:
                changed[name] = new_text.encode("utf-8")

        # rich strings that name the font explicitly -> fix in place
        out_si = list(si_orig)
        for i, inner in enumerate(si_inner):
            if from_l in inner.lower():
                r = build_si(inner, None, from_l, to_font, khmer_font)
                if r is not None:
                    out_si[i] = f"<si>{r}</si>"
        if sst and (new_si or out_si != si_orig):
            total = len(out_si) + len(new_si)
            hdr = re.sub(r'uniqueCount="\d+"', f'uniqueCount="{total}"', header)
            changed["xl/sharedStrings.xml"] = (hdr + "".join(out_si) + "".join(new_si) + footer).encode("utf-8")

        # text boxes, comments, charts
        fpat = re.compile(r'(?<=")' + re.escape(from_font) + r'(?=")', re.I)
        for name in names:
            if name.startswith(("xl/drawings/", "xl/comments", "xl/charts/")) and name.endswith(".xml"):
                raw = zin.read(name).decode("utf-8")
                if from_l in raw.lower():
                    changed[name] = fpat.sub(_esc_attr(to_font), raw).encode("utf-8")

        with zipfile.ZipFile(dst, "w", zipfile.ZIP_STORED) as zout:  # temp file: no compression
            for info in zin.infolist():
                data = changed.get(info.filename)
                if data is None:
                    data = zin.read(info.filename)
                zi = zipfile.ZipInfo(info.filename, info.date_time)
                zi.compress_type = zipfile.ZIP_STORED
                zi.external_attr = info.external_attr
                zout.writestr(zi, data)
    return stats
