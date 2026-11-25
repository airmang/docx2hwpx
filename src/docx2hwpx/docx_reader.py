"""Read a .docx (zipfile + lxml) into the intermediate model of model.py.

Everything Word shows is either carried into the model or counted in the LossLog.
Tracked changes are accepted (insertions kept, deletions dropped).
"""
from __future__ import annotations

import posixpath
import re
import zipfile
from dataclasses import dataclass, field

from lxml import etree

from . import model as M
from .losses import LossLog
from .styles import W, W_NS, StyleSheet, Theme, merge, on, parse_borders, parse_margins, parse_ppr, parse_rpr, \
    to_int, wval

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
R = "{%s}" % R_NS
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
PIC_URI = "http://schemas.openxmlformats.org/drawingml/2006/picture"
CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DGM_URI = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
WPS_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
WPG_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
WPC_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
V = "{urn:schemas-microsoft-com:vml}"
MATH = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
PKG_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"

# Strict OOXML namespaces -> transitional, so one reader handles both
STRICT = {
    b"http://purl.oclc.org/ooxml/wordprocessingml/main": W_NS.encode(),
    b"http://purl.oclc.org/ooxml/officeDocument/relationships": R_NS.encode(),
    b"http://purl.oclc.org/ooxml/drawingml/main": b"http://schemas.openxmlformats.org/drawingml/2006/main",
    b"http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing":
        b"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    b"http://purl.oclc.org/ooxml/drawingml/picture": PIC_URI.encode(),
    b"http://purl.oclc.org/ooxml/drawingml/chart": CHART_URI.encode(),
}

IMAGE_EXT = {"png": "png", "jpg": "jpg", "jpeg": "jpg", "jpe": "jpg", "gif": "gif", "bmp": "bmp", "dib": "bmp",
             "tif": "tif", "tiff": "tif", "emf": "emf", "wmf": "wmf", "svg": "svg"}

# Symbol / Wingdings private-use bullet characters -> Unicode
SYMBOL_MAP = {
    0xF0B7: "\u2022", 0xF0A7: "\u25AA", 0xF0D8: "\u27A2", 0xF076: "\u2756", 0xF0FC: "\u2713", 0xF0A8: "\u25A1",
    0xF06E: "\u25A0", 0xF06C: "\u25CF", 0xF0E0: "\u2192", 0xF0D7: "\u25B8", 0xF02D: "-", 0xF0B2: "\u25C7",
    0xF071: "\u2752", 0xF075: "\u25C6", 0xF0A1: "\u25CB", 0xF09F: "\u2022", 0xF0FB: "\u2717", 0xF0E8: "\u2794",
}


class DocxError(Exception):
    """The file is not a readable .docx; message is Korean, one line."""


def _clean_char(ch: str) -> str:
    cp = ord(ch)
    if cp in SYMBOL_MAP:
        return SYMBOL_MAP[cp]
    if 0xF000 <= cp <= 0xF0FF:
        return "\u2022"
    return ch


def map_symbol_text(s: str) -> str:
    return "".join(_clean_char(c) for c in s)


@dataclass
class Ctx:
    """Where a block is read: which part (for rels) and which table style / conditional formats apply."""
    part: str
    table_ppr: dict = field(default_factory=dict)
    table_rpr: dict = field(default_factory=dict)


@dataclass
class Story:
    """State that spans paragraphs inside one story (fields, carried page breaks)."""
    fields: list = field(default_factory=list)
    carry_break: str | None = None


class DocxReader:
    def __init__(self, data: bytes | str, losses: LossLog | None = None):
        try:
            self.zip = zipfile.ZipFile(data if isinstance(data, str) else __import__("io").BytesIO(data))
        except zipfile.BadZipFile as ex:
            raise DocxError("파일이 손상되었거나 Word 문서(.docx)가 아닙니다.") from ex
        self.names = set(self.zip.namelist())
        self.losses = losses or LossLog()
        self._rels_cache: dict[str, dict] = {}
        self._xml_cache: dict[str, etree._Element] = {}
        self.main = self._main_part()
        rels = self.rels(self.main)
        theme_part = self._rel_part(rels, "/theme")
        self.theme = Theme(self._read(theme_part) if theme_part else None)
        styles_part = self._rel_part(rels, "/styles")
        self.styles = StyleSheet(self._read(styles_part) if styles_part else None, self.theme)
        num_part = self._rel_part(rels, "/numbering")
        self.numbering = self._parse_numbering(num_part)
        settings_part = self._rel_part(rels, "/settings")
        self.even_odd = False
        if settings_part:
            st = self.xml(settings_part)
            if st is not None and st.find(W + "evenAndOddHeaders") is not None:
                self.even_odd = bool(on(st.find(W + "evenAndOddHeaders")))
        self.notes = {"footnote": {}, "endnote": {}}
        for kind in ("footnote", "endnote"):
            part = self._rel_part(rels, "/" + kind + "s")
            if part:
                root = self.xml(part)
                if root is not None:
                    for n in root.findall(W + kind):
                        if n.get(W + "type") in (None, "normal"):
                            self.notes[kind][n.get(W + "id")] = (part, n)
        self._hf_cache: dict[str, list] = {}

    # ------------------------------------------------------------------ package
    def _read(self, name: str | None) -> bytes | None:
        if not name or name not in self.names:
            return None
        data = self.zip.read(name)
        for k, v in STRICT.items():
            if k in data:
                data = data.replace(k, v)
        return data

    def xml(self, name: str):
        if name not in self._xml_cache:
            data = self._read(name)
            if data is None:
                self._xml_cache[name] = None
            else:
                try:
                    self._xml_cache[name] = etree.fromstring(data, etree.XMLParser(huge_tree=True, recover=False))
                except etree.XMLSyntaxError as ex:
                    raise DocxError("문서 내용(XML)이 손상되어 읽을 수 없습니다.") from ex
        return self._xml_cache[name]

    def _main_part(self) -> str:
        if "_rels/.rels" in self.names:
            root = etree.fromstring(self.zip.read("_rels/.rels"))
            for rel in root:
                typ = rel.get("Type") or ""
                if typ.endswith("/officeDocument"):
                    target = (rel.get("Target") or "").lstrip("/")
                    if target in self.names:
                        return target
        if "word/document.xml" in self.names:
            return "word/document.xml"
        raise DocxError("Word 문서(.docx)가 아닙니다.")

    @staticmethod
    def _rel_part(rels: dict, suffix: str) -> str | None:
        for key, val in rels.items():
            if not key.startswith("#") and val[0].endswith(suffix) and val[1] != "External":
                return key
        return None

    def rels(self, part: str) -> dict:
        """target part (resolved) or external URL -> (type, mode); keyed by rId under '#id' too."""
        if part in self._rels_cache:
            return self._rels_cache[part]
        d, base = posixpath.split(part)
        rel_name = posixpath.join(d, "_rels", base + ".rels")
        out: dict = {}
        if rel_name in self.names:
            root = etree.fromstring(self.zip.read(rel_name))
            for rel in root:
                typ = rel.get("Type") or ""
                target = rel.get("Target") or ""
                mode = rel.get("TargetMode") or "Internal"
                if mode != "External":
                    target = target.lstrip("/") if target.startswith("/") else posixpath.normpath(
                        posixpath.join(d, target))
                out["#" + (rel.get("Id") or "")] = (target, typ, mode)
                out[target] = (typ, mode)
        self._rels_cache[part] = out
        return out

    def rel_target(self, part: str, rid: str | None):
        if not rid:
            return None
        return self.rels(part).get("#" + rid)

    # ------------------------------------------------------------------ numbering
    def _parse_numbering(self, part: str | None) -> dict:
        out: dict = {}
        if not part:
            return out
        root = self.xml(part)
        if root is None:
            return out
        abstract = {}
        for an in root.findall(W + "abstractNum"):
            levels = {}
            for lvl in an.findall(W + "lvl"):
                ilvl = to_int(lvl.get(W + "ilvl"), 0)
                text_el = lvl.find(W + "lvlText")
                rpr = parse_rpr(lvl.find(W + "rPr"), self.theme)
                levels[ilvl] = {
                    "start": to_int(wval(lvl.find(W + "start")), 1),
                    "fmt": wval(lvl.find(W + "numFmt"), default="decimal"),
                    "text": wval(text_el, default="") if text_el is not None else "",
                    "ppr": parse_ppr(lvl.find(W + "pPr")),
                    "font": rpr.get("font_ascii") or rpr.get("font_eastAsia"),
                }
            link = wval(an.find(W + "numStyleLink"))
            abstract[an.get(W + "abstractNumId")] = (levels, link)
        for num in root.findall(W + "num"):
            nid = num.get(W + "numId")
            aid = wval(num.find(W + "abstractNumId"))
            levels, link = abstract.get(aid, ({}, None))
            if link and not levels:
                # numStyleLink: the numbering style names the num that holds the levels
                st = self.styles.styles.get(link)
                if st is not None:
                    snid = wval(st.find(f"{W}pPr/{W}numPr/{W}numId"))
                    snum = root.find(f"{W}num[@{W}numId='{snid}']") if snid else None
                    if snum is not None:
                        levels = abstract.get(wval(snum.find(W + "abstractNumId")), ({}, None))[0]
            levels = {k: dict(v) for k, v in levels.items()}
            for ov in num.findall(W + "lvlOverride"):
                ilvl = to_int(ov.get(W + "ilvl"), 0)
                so = ov.find(W + "startOverride")
                lvl_el = ov.find(W + "lvl")
                if lvl_el is not None:
                    text_el = lvl_el.find(W + "lvlText")
                    levels[ilvl] = {
                        "start": to_int(wval(lvl_el.find(W + "start")), 1),
                        "fmt": wval(lvl_el.find(W + "numFmt"), default="decimal"),
                        "text": wval(text_el, default="") if text_el is not None else "",
                        "ppr": parse_ppr(lvl_el.find(W + "pPr")),
                        "font": None,
                    }
                if so is not None and ilvl in levels:
                    levels[ilvl]["start"] = to_int(wval(so), 1)
            out[nid] = levels
        return out

    # ------------------------------------------------------------------ document
    def read(self) -> M.Document:
        root = self.xml(self.main)
        if root is None:
            raise DocxError("Word 문서(.docx)가 아닙니다.")
        body = root.find(W + "body")
        if body is None:
            raise DocxError("Word 문서의 본문을 찾을 수 없습니다.")
        sections: list[M.Section] = []
        cur: list = []
        story = Story()
        ctx = Ctx(self.main)
        for el in self._block_children(body):
            tag = etree.QName(el).localname
            if tag == "sectPr":
                continue
            if tag == "p":
                blocks, sect = self.paragraph(el, ctx, story)
                cur.extend(blocks)
                if sect is not None:
                    sec = self.section(sect, sections[-1] if sections else None)
                    sec.blocks = cur
                    sections.append(sec)
                    cur = []
            elif tag == "tbl":
                cur.append(self.table(el, ctx, story))
            elif tag == "altChunk":
                self.losses.add("alt_chunk")
        final = body.find(W + "sectPr")
        sec = self.section(final, sections[-1] if sections else None)
        sec.blocks = cur
        sections.append(sec)
        return M.Document(sections=sections, even_odd_headers=self.even_odd, numbering=self.numbering)

    def _block_children(self, parent):
        """Block-level children with content controls and customXml unwrapped."""
        for el in parent:
            if not isinstance(el.tag, str):
                continue
            tag = etree.QName(el).localname
            ns = etree.QName(el).namespace
            if ns == W_NS and tag == "sdt":
                content = el.find(W + "sdtContent")
                if content is not None:
                    yield from self._block_children(content)
            elif ns == W_NS and tag in ("customXml", "ins", "moveTo"):
                if tag != "customXml":
                    self.losses.add("tracked_change")
                yield from self._block_children(el)
            elif ns == W_NS and tag in ("del", "moveFrom"):
                self.losses.add("tracked_change")
            elif tag == "AlternateContent":
                yield from self._block_children(self._alt_choice(el))
            elif ns == W_NS:
                yield el

    def blocks(self, parent, ctx: Ctx, story: Story | None = None) -> list:
        story = story or Story()
        out: list = []
        for el in self._block_children(parent):
            tag = etree.QName(el).localname
            if tag == "p":
                blocks, _sect = self.paragraph(el, ctx, story)
                out.extend(blocks)
            elif tag == "tbl":
                out.append(self.table(el, ctx, story))
            elif tag == "altChunk":
                self.losses.add("alt_chunk")
        return out

    # ------------------------------------------------------------------ sections
    def section(self, sp, prev: M.Section | None) -> M.Section:
        sec = M.Section()
        if prev is not None:
            # header/footer references are inherited from the previous section
            sec.headers = M.HeaderFooterSet(prev.headers.default, prev.headers.first, prev.headers.even)
            sec.footers = M.HeaderFooterSet(prev.footers.default, prev.footers.first, prev.footers.even)
        if sp is None:
            return sec
        sz = sp.find(W + "pgSz")
        if sz is not None:
            sec.page_w = to_int(sz.get(W + "w"), 11906)
            sec.page_h = to_int(sz.get(W + "h"), 16838)
            sec.landscape = sz.get(W + "orient") == "landscape" or sec.page_w > sec.page_h
        mar = sp.find(W + "pgMar")
        if mar is not None:
            for k in ("top", "bottom", "left", "right", "header", "footer", "gutter"):
                v = mar.get(W + k)
                if v is not None:
                    sec.margins[k] = to_int(v, 0)
        cols = sp.find(W + "cols")
        if cols is not None:
            sec.cols = max(1, to_int(cols.get(W + "num"), 1))
            sec.col_space = to_int(cols.get(W + "space"), 720)
            sec.col_sep = cols.get(W + "sep") in ("1", "true", "on")
            if cols.get(W + "equalWidth") in ("0", "false", "off"):
                widths = [(to_int(c.get(W + "w"), 0), to_int(c.get(W + "space"), 0)) for c in cols.findall(W + "col")]
                if widths:
                    sec.col_widths = widths
                    sec.cols = len(widths)
        sec.start_type = wval(sp.find(W + "type"), default="nextPage")
        sec.title_pg = bool(on(sp.find(W + "titlePg")))
        pn = sp.find(W + "pgNumType")
        if pn is not None:
            if pn.get(W + "start") is not None:
                sec.pg_start = to_int(pn.get(W + "start"), 1)
            sec.pg_fmt = pn.get(W + "fmt")
        if sp.find(W + "pgBorders") is not None:
            self.losses.add("page_border")
        if sp.find(W + "lnNumType") is not None:
            self.losses.add("line_numbers")
        for tag, target in (("headerReference", sec.headers), ("footerReference", sec.footers)):
            for ref in sp.findall(W + tag):
                kind = ref.get(W + "type") or "default"
                blocks = self.header_footer(ref.get(R + "id"))
                if kind in ("default", "first", "even"):
                    setattr(target, kind, blocks)
        return sec

    def header_footer(self, rid: str | None) -> list:
        rel = self.rel_target(self.main, rid)
        if not rel:
            return []
        part = rel[0]
        if part in self._hf_cache:
            return self._hf_cache[part]
        root = self.xml(part)
        blocks = self.blocks(root, Ctx(part)) if root is not None else []
        self._hf_cache[part] = blocks
        return blocks

    # ------------------------------------------------------------------ paragraphs
    def para_props(self, p_el, ctx: Ctx) -> tuple[dict, dict, dict, str | None]:
        """Resolved (ppr, base rpr for runs, paragraph-mark rpr, style id)."""
        direct = parse_ppr(p_el.find(W + "pPr")) if p_el is not None else {}
        sid = direct.get("pStyle") or self.styles.default_para
        sppr, srpr = self.styles.para_style(sid)
        ss = self.styles
        if sid == ss.default_para or sid not in ss.styles:
            # Normal-styled text inside a table: the table style wins over Normal
            ppr = merge(merge(merge(ss.doc_ppr, sppr), ctx.table_ppr), {})
            rpr = merge(merge(ss.doc_rpr, srpr), ctx.table_rpr)
        else:
            ppr = merge(merge(ss.doc_ppr, ctx.table_ppr), sppr)
            rpr = merge(merge(ss.doc_rpr, ctx.table_rpr), srpr)
        # numbering level indentation sits between the style and direct formatting
        nid = direct.get("numId", ppr.get("numId"))
        ilvl = direct.get("ilvl", ppr.get("ilvl", 0))
        if nid and nid != "0" and nid in self.numbering:
            lvl = self.numbering[nid].get(ilvl)
            if lvl:
                ppr = merge(ppr, {k: v for k, v in lvl["ppr"].items() if k in ("left", "right", "first", "tabs")})
        ppr = merge(ppr, direct)
        if direct.get("numId") is not None or "ilvl" in direct:
            ppr["numId"] = nid
            ppr["ilvl"] = ilvl
        mark = parse_rpr(p_el.find(f"{W}pPr/{W}rPr"), self.theme) if p_el is not None else {}
        mark_rpr = merge(rpr, self._char_style_rpr(mark))
        return ppr, rpr, mark_rpr, sid

    def _char_style_rpr(self, direct: dict) -> dict:
        rs = direct.get("rStyle")
        if rs:
            return merge(self.styles.char_style(rs), direct)
        return direct

    def paragraph(self, p_el, ctx: Ctx, story: Story):
        ppr, base_rpr, mark_rpr, sid = self.para_props(p_el, ctx)
        sect = p_el.find(f"{W}pPr/{W}sectPr")
        segments: list[tuple[str | None, list]] = [(None, [])]
        after: list = []
        state = {"segments": segments, "after": after, "base": base_rpr, "ctx": ctx, "story": story}
        self._inline_children(p_el, state, target=None)
        num = None
        nid = ppr.get("numId")
        if nid and nid != "0" and nid in self.numbering:
            num = (nid, int(ppr.get("ilvl") or 0))
        out: list = []
        for i, (brk, items) in enumerate(segments):
            last = i == len(segments) - 1
            if i > 0 and last and not items and not (after and last):
                # a page break ending the paragraph: the next paragraph starts the new page
                story.carry_break = brk
                break
            para = M.Paragraph(ppr=ppr, mark_rpr=mark_rpr, items=items, style_id=sid, num=num if i == 0 else None)
            if i == 0 and story.carry_break and not (sect is not None and not items):
                setattr(para, "page_break" if story.carry_break == "page" else "column_break", True)
                story.carry_break = None
            if i > 0:
                setattr(para, "page_break" if brk == "page" else "column_break", True)
            out.append(para)
        if out:
            out[-1].after = after
        if sect is not None and out and all(p.is_empty() for p in out if isinstance(p, M.Paragraph)):
            # the empty paragraph that only carries a section break is not a line in Hangul
            if not any(p.page_break or p.column_break for p in out):
                out = []
        return out, sect

    # ------------------------------------------------------------------ inline content
    def _emit(self, state, item, target):
        if target is not None:
            target.append(item)
            return
        story: Story = state["story"]
        for f in reversed(story.fields):
            if f.get("collect") is not None and f["phase"] == "result":
                f["collect"].append(item)
                return
            if f["phase"] == "result" and f.get("skip"):
                return
        state["segments"][-1][1].append(item)

    def _skipping(self, state) -> bool:
        story: Story = state["story"]
        for f in story.fields:
            if f["phase"] == "instr":
                return True
            if f["phase"] == "result" and f.get("skip"):
                return True
        return False

    def _alt_choice(self, el):
        choice = el.find(MC + "Choice")
        if choice is not None:
            req = choice.get("Requires") or ""
            if any(r in req for r in ("wps", "wpg", "wpc", "w14", "w15", "v", "a14", "wp14", "cx", "c14")):
                return choice
        fb = el.find(MC + "Fallback")
        return fb if fb is not None else (choice if choice is not None else el)

    def _inline_children(self, parent, state, target):
        for el in parent:
            if not isinstance(el.tag, str):
                continue
            q = etree.QName(el)
            tag, ns = q.localname, q.namespace
            if ns == W_NS:
                if tag == "r":
                    self._run(el, state, target)
                elif tag == "hyperlink":
                    self._hyperlink(el, state, target)
                elif tag == "fldSimple":
                    self._fld_simple(el, state, target)
                elif tag in ("ins", "moveTo"):
                    self.losses.add("tracked_change")
                    self._inline_children(el, state, target)
                elif tag in ("del", "moveFrom"):
                    self.losses.add("tracked_change")
                elif tag == "sdt":
                    content = el.find(W + "sdtContent")
                    if content is not None:
                        self._inline_children(content, state, target)
                elif tag in ("smartTag", "customXml", "dir", "bdo"):
                    self._inline_children(el, state, target)
                elif tag in ("commentRangeStart",):
                    self.losses.add("comment")
            elif ns == MATH[1:-1]:
                if tag in ("oMath", "oMathPara"):
                    self._math(el, state, target)
            elif tag == "AlternateContent":
                self._inline_children(self._alt_choice(el), state, target)

    def _run_rpr(self, r_el, state) -> dict:
        direct = parse_rpr(r_el.find(W + "rPr"), self.theme)
        return merge(state["base"], self._char_style_rpr(direct))

    def _run(self, r_el, state, target):
        rpr = self._run_rpr(r_el, state)
        story: Story = state["story"]
        buf: list[str] = []

        def flush():
            if buf:
                text = "".join(buf)
                buf.clear()
                if rpr.get("vanish"):
                    self.losses.add("hidden_text")
                    return
                if rpr.get("caps") or rpr.get("smallCaps"):
                    text = text.upper()
                if self._skipping(state):
                    return
                self._emit(state, M.Text(text, rpr), target)

        for el in r_el:
            if not isinstance(el.tag, str):
                continue
            q = etree.QName(el)
            tag, ns = q.localname, q.namespace
            if ns == W_NS:
                if tag == "t":
                    buf.append(map_symbol_text(el.text or "") if rpr.get("font_ascii") in ("Symbol", "Wingdings")
                               else (el.text or ""))
                elif tag in ("tab", "ptab"):
                    buf.append("\t")
                elif tag in ("br", "cr"):
                    kind = el.get(W + "type") or "textWrapping"
                    if kind in ("page", "column") and target is None and not self._skipping(state):
                        flush()
                        state["segments"].append((kind, []))
                    else:
                        buf.append("\n")
                elif tag == "noBreakHyphen":
                    buf.append("\u2011")
                elif tag == "sym":
                    ch = el.get(W + "char") or ""
                    try:
                        buf.append(_clean_char(chr(int(ch, 16))))
                    except ValueError:
                        pass
                elif tag == "instrText":
                    if story.fields and story.fields[-1]["phase"] == "instr":
                        story.fields[-1]["instr"] += el.text or ""
                elif tag == "fldChar":
                    flush()
                    self._fld_char(el, rpr, state, target)
                elif tag in ("drawing", "pict", "object"):
                    flush()
                    if not self._skipping(state):
                        self._object(el, rpr, state, target)
                elif tag in ("footnoteReference", "endnoteReference"):
                    flush()
                    kind = "footnote" if tag.startswith("foot") else "endnote"
                    self._note_ref(kind, el.get(W + "id"), rpr, state, target)
                elif tag == "commentReference":
                    pass  # counted at commentRangeStart; a reference without a range is counted here
                elif tag == "pgNum":
                    flush()
                    self._emit(state, M.Field("PAGE", rpr), target)
                elif tag == "ruby":
                    base = el.find(f"{W}rubyBase")
                    if base is not None:
                        flush()
                        self._inline_children(base, dict(state), target)
                elif tag == "delText":
                    pass
            elif tag == "AlternateContent":
                flush()
                choice = self._alt_choice(el)
                fake = etree.Element(W + "r")
                for c in choice:
                    fake.append(c.__copy__())
                if r_el.find(W + "rPr") is not None:
                    fake.insert(0, r_el.find(W + "rPr").__copy__())
                self._run(fake, state, target)
        flush()
        if r_el.find(W + "commentReference") is not None and not self._has_range(r_el):
            self.losses.add("comment")

    def _has_range(self, r_el) -> bool:
        p = r_el.getparent()
        while p is not None and etree.QName(p).localname != "p":
            p = p.getparent()
        if p is None:
            return False
        cid = r_el.find(W + "commentReference").get(W + "id")
        return p.getroottree().getroot().find(f".//{W}commentRangeStart[@{W}id='{cid}']") is not None

    # ---------------------------------------------------------------- fields
    def _fld_char(self, el, rpr, state, target):
        story: Story = state["story"]
        kind = el.get(W + "fldCharType")
        if kind == "begin":
            story.fields.append({"phase": "instr", "instr": "", "rpr": rpr, "collect": None, "skip": False})
        elif kind == "separate" and story.fields:
            f = story.fields[-1]
            self._field_start(f, state, target)
        elif kind == "end" and story.fields:
            f = story.fields[-1]
            if f["phase"] == "instr":
                self._field_start(f, state, target)
            story.fields.pop()
            if f.get("hyperlink") is not None:
                self._emit(state, f["hyperlink"], target)

    def _field_start(self, f, state, target):
        instr = f["instr"].strip()
        word = instr.split(" ", 1)[0].upper() if instr else ""
        f["phase"] = "result"
        if word in ("PAGE", "NUMPAGES", "SECTIONPAGES"):
            f["skip"] = True
            fmt = "DIGIT"
            m = re.search(r"\\\*\s*(\w+)", instr)
            if m:
                fmt = {"roman": "ROMAN_SMALL", "ROMAN": "ROMAN_CAPITAL", "alphabetic": "LATIN_SMALL",
                       "ALPHABETIC": "LATIN_CAPITAL"}.get(m.group(1), "DIGIT")
            self._emit_raw(state, M.Field("PAGE" if word == "PAGE" else "NUMPAGES", f["rpr"], fmt), target)
        elif word == "HYPERLINK":
            url = _hyperlink_url(instr)
            items: list = []
            f["collect"] = items
            f["hyperlink"] = M.Hyperlink(url, items)
            if url is None:
                self.losses.add("internal_link")
        elif word in ("", "MERGEFORMAT"):
            pass
        elif word in ("FORMTEXT", "FORMCHECKBOX", "FORMDROPDOWN"):
            self.losses.add("form_control")
        elif word not in ("EQ",) and len(state["story"].fields) == 1:
            self.losses.add("field")  # fields nested in a field (page refs of a TOC) are part of it

    def _emit_raw(self, state, item, target):
        """Emit while the field itself is in its result phase (bypassing its own skip)."""
        if target is not None:
            target.append(item)
            return
        story: Story = state["story"]
        for f in reversed(story.fields[:-1]):
            if f.get("collect") is not None and f["phase"] == "result":
                f["collect"].append(item)
                return
        state["segments"][-1][1].append(item)

    def _fld_simple(self, el, state, target):
        instr = (el.get(W + "instr") or "").strip()
        word = instr.split(" ", 1)[0].upper() if instr else ""
        first_r = el.find(W + "r")
        rpr = self._run_rpr(first_r, state) if first_r is not None else dict(state["base"])
        if word in ("PAGE", "NUMPAGES", "SECTIONPAGES"):
            self._emit(state, M.Field("PAGE" if word == "PAGE" else "NUMPAGES", rpr), target)
            return
        if word == "HYPERLINK":
            items: list = []
            self._inline_children(el, state, items)
            url = _hyperlink_url(instr)
            if url is None:
                self.losses.add("internal_link")
            self._emit(state, M.Hyperlink(url, items), target)
            return
        if word:
            self.losses.add("field")
        self._inline_children(el, state, target)

    def _hyperlink(self, el, state, target):
        rid = el.get(R + "id")
        url = None
        if rid:
            rel = self.rel_target(state["ctx"].part, rid)
            if rel and rel[2] == "External":
                url = rel[0]
                anchor = el.get(W + "anchor")
                if anchor:
                    url += "#" + anchor
        if url is None and not state["story"].fields:
            self.losses.add("internal_link")  # (a TOC's entry links belong to the TOC field)
        items: list = []
        self._inline_children(el, state, items)
        if self._skipping(state):
            return
        self._emit(state, M.Hyperlink(url, items), target)

    def _note_ref(self, kind, nid, rpr, state, target):
        entry = self.notes[kind].get(nid)
        blocks: list = []
        if entry is not None:
            part, el = entry
            blocks = self.blocks(el, Ctx(part))
        self._emit(state, M.NoteRef(kind, blocks, rpr), target)

    def _math(self, el, state, target):
        text = "".join(t.text or "" for t in el.iter(MATH + "t"))
        self.losses.add("equation")
        if text:
            self._emit(state, M.Text(text, dict(state["base"])), target)

    # ---------------------------------------------------------------- drawings
    @staticmethod
    def _anchor(holder) -> dict:
        """Placement of a floating drawing: wrap, relative-to frames, alignment or offsets."""
        out = {"wrap": "square", "behind": holder.get("behindDoc") in ("1", "true"),
               "overlap": holder.get("allowOverlap") not in ("0", "false")}
        for tag in ("wrapSquare", "wrapTight", "wrapThrough", "wrapTopAndBottom", "wrapNone"):
            if holder.find(WP + tag) is not None:
                out["wrap"] = tag[4:].lower()
                break
        for axis, tag in (("h", "positionH"), ("v", "positionV")):
            el = holder.find(WP + tag)
            if el is None:
                continue
            out[axis + "_rel"] = el.get("relativeFrom") or ""
            al = el.find(WP + "align")
            off = el.find(WP + "posOffset")
            if al is not None and al.text:
                out[axis + "_align"] = al.text.strip()
            elif off is not None and off.text:
                out[axis + "_off"] = to_int(off.text.strip(), 0)
        for k in ("distT", "distB", "distL", "distR"):
            out[k] = to_int(holder.get(k), 0)
        return out

    def _image(self, part: str, rid: str | None, cx: int, cy: int, rpr: dict, floating: bool):
        rel = self.rel_target(part, rid)
        if not rel or rel[2] == "External":
            self.losses.add("linked_image")
            return None
        name = rel[0]
        ext = name.rsplit(".", 1)[-1].lower()
        fmt = IMAGE_EXT.get(ext)
        data = self.zip.read(name) if name in self.names else None
        if not fmt or not data:
            self.losses.add("image_unsupported")
            return None
        return M.Image(data, fmt, max(cx, 1), max(cy, 1), rpr, floating)

    def _object(self, el, rpr, state, target):
        ctx: Ctx = state["ctx"]
        tag = etree.QName(el).localname
        if tag == "drawing":
            for holder in el:
                if not isinstance(holder.tag, str):
                    continue
                floating = etree.QName(holder).localname == "anchor"
                ext = holder.find(WP + "extent")
                cx = to_int(ext.get("cx"), 0) if ext is not None else 0
                cy = to_int(ext.get("cy"), 0) if ext is not None else 0
                gd = holder.find(f"{A}graphic/{A}graphicData")
                uri = gd.get("uri") if gd is not None else ""
                if uri == PIC_URI:
                    blip = gd.find(f".//{A}blip")
                    rid = blip.get(R + "embed") if blip is not None else None
                    if blip is not None and not rid and blip.get(R + "link"):
                        self.losses.add("linked_image")
                        continue
                    img = self._image(ctx.part, rid, cx, cy, rpr, floating)
                    if img is not None:
                        if floating:
                            img.anchor = self._anchor(holder)
                        self._emit(state, img, target)
                elif uri == WPS_URI:
                    txbx = gd.find(f".//{W}txbxContent")
                    if txbx is not None:
                        self._textbox(txbx, state)
                    else:
                        self.losses.add("shape")
                elif uri in (WPG_URI, WPC_URI):
                    found = False
                    for txbx in gd.iter(W + "txbxContent"):
                        self._textbox(txbx, state)
                        found = True
                    if not found:
                        self.losses.add("shape")
                elif uri == CHART_URI:
                    self.losses.add("chart")
                elif uri == DGM_URI:
                    self.losses.add("smartart")
                else:
                    self.losses.add("shape")
        else:  # VML w:pict / w:object
            imagedata = el.find(f".//{V}imagedata")
            txbx = el.find(f".//{W}txbxContent")
            if txbx is not None:
                self._textbox(txbx, state)
            elif imagedata is not None:
                shape = imagedata.getparent()
                cx, cy = _vml_size(shape.get("style") or "")
                floating = "position:absolute" in (shape.get("style") or "").replace(" ", "")
                if tag == "object":
                    self.losses.add("ole_object")
                img = self._image(ctx.part, imagedata.get(R + "id"), cx, cy, rpr, floating)
                if img is not None:
                    if floating:
                        self.losses.add("floating_picture")  # VML placement: kept in the line
                        img.floating = False
                    self._emit(state, img, target)
            elif tag == "object":
                self.losses.add("ole_dropped")
            elif el.find(f".//{V}line") is not None or el.find(f".//{V}rect") is not None or len(el):
                self.losses.add("shape")

    def _textbox(self, txbx, state):
        self.losses.add("textbox")
        state["after"].extend(self.blocks(txbx, state["ctx"]))

    # ------------------------------------------------------------------ tables
    def table(self, tbl, ctx: Ctx, story: Story) -> M.Table:
        tblpr = tbl.find(W + "tblPr")
        sid = wval(tblpr.find(W + "tblStyle")) if tblpr is not None else None
        ts = self.styles.table_style(sid)
        borders = dict(ts["borders"])
        margins = {"left": 108, "right": 108, "top": 0, "bottom": 0}
        margins.update(ts["margins"])
        width = None
        jc = None
        indent = 0
        look = {"firstRow": True, "lastRow": False, "firstColumn": True, "lastColumn": False,
                "noHBand": False, "noVBand": True}
        if tblpr is not None:
            borders.update(parse_borders(tblpr.find(W + "tblBorders")))
            margins.update(parse_margins(tblpr.find(W + "tblCellMar")))
            tw = tblpr.find(W + "tblW")
            if tw is not None:
                width = (to_int(tw.get(W + "w"), 0), tw.get(W + "type") or "dxa")
            jc = wval(tblpr.find(W + "jc"))
            ind = tblpr.find(W + "tblInd")
            if ind is not None and (ind.get(W + "type") in (None, "dxa")):
                indent = to_int(ind.get(W + "w"), 0)
            lk = tblpr.find(W + "tblLook")
            if lk is not None:
                v = lk.get(W + "val")
                if v and all(lk.get(W + k) is None for k in look):
                    bits = int(v, 16)
                    look = {"firstRow": bool(bits & 0x20), "lastRow": bool(bits & 0x40),
                            "firstColumn": bool(bits & 0x80), "lastColumn": bool(bits & 0x100),
                            "noHBand": bool(bits & 0x200), "noVBand": bool(bits & 0x400)}
                for k in look:
                    if lk.get(W + k) is not None:
                        look[k] = lk.get(W + k) in ("1", "true", "on")
        grid = [to_int(g.get(W + "w"), 0) for g in tbl.findall(f"{W}tblGrid/{W}gridCol")]
        # rows (content controls around rows/cells unwrapped)
        tr_els = [el for el in self._block_children(tbl) if etree.QName(el).localname == "tr"]
        nrows = len(tr_els)
        rows: list[M.Row] = []
        for ri, tr in enumerate(tr_els):
            trpr = tr.find(W + "trPr")
            row = M.Row(cells=[])
            grid_before = 0
            if trpr is not None:
                th = trpr.find(W + "trHeight")
                if th is not None:
                    row.height = to_int(th.get(W + "val"), 0)
                    row.height_rule = th.get(W + "hRule") or "atLeast"
                row.header = bool(on(trpr.find(W + "tblHeader")))
                grid_before = to_int(wval(trpr.find(W + "gridBefore")), 0)
            tc_els = [el for el in self._block_children(tr) if etree.QName(el).localname == "tc"]
            col = 0
            for _ in range(grid_before):
                row.cells.append(M.Cell(blocks=[], borders={s: None for s in ("top", "left", "bottom", "right")}))
                col += 1
            ncols = len(grid)
            for tc in tc_els:
                tcpr = tc.find(W + "tcPr")
                span = 1
                vmerge = None
                cw = None
                cborders = {}
                fill = None
                valign = None
                cmar = {}
                if tcpr is not None:
                    span = max(1, to_int(wval(tcpr.find(W + "gridSpan")), 1))
                    vm = tcpr.find(W + "vMerge")
                    if vm is not None:
                        vmerge = vm.get(W + "val") or "continue"
                    tw = tcpr.find(W + "tcW")
                    if tw is not None and (tw.get(W + "type") in (None, "dxa")):
                        cw = to_int(tw.get(W + "w"), 0) or None
                    cborders = parse_borders(tcpr.find(W + "tcBorders"))
                    shd = tcpr.find(W + "shd")
                    if shd is not None:
                        f_ = (shd.get(W + "fill") or "auto")
                        fill = None if f_.lower() == "auto" else f_
                        if fill is None and shd.get(W + "val") == "solid" and (shd.get(W + "color") or "auto") != "auto":
                            fill = shd.get(W + "color")
                    valign = wval(tcpr.find(W + "vAlign"))
                    cmar = parse_margins(tcpr.find(W + "tcMar"))
                    if tcpr.find(W + "textDirection") is not None and \
                            wval(tcpr.find(W + "textDirection")) not in ("lrTb", "tb"):
                        self.losses.add("vertical_text")
                # conditional formatting of the table style for this cell
                cond = self._conditions(ts, look, ri, nrows, col, span, ncols, rows_header=row.header)
                cppr, crpr, cbord, cshd = {}, {}, {}, ts["shd"]
                for c in cond:
                    cppr = merge(cppr, c["ppr"])
                    crpr = merge(crpr, c["rpr"])
                    cbord.update(c["borders"])
                    if c["shd"]:
                        cshd = c["shd"]
                if fill is None and tcpr is not None and tcpr.find(W + "shd") is not None:
                    cshd = None
                cell_ctx = Ctx(ctx.part, merge(ts["ppr"], cppr), merge(ts["rpr"], crpr))
                blocks = self.blocks(tc, cell_ctx, Story())
                eff_borders = self._cell_borders(borders, cbord, cborders, ri, nrows, col, span, ncols)
                row.cells.append(M.Cell(blocks=blocks, span=span, vmerge=vmerge, width=cw, borders=eff_borders,
                                        fill=fill or cshd, valign=valign, margins=cmar))
                col += span
            rows.append(row)
        return M.Table(grid=grid, rows=rows, width=width, jc=jc, indent=indent, cell_margins=margins)

    @staticmethod
    def _conditions(ts, look, ri, nrows, col, span, ncols, rows_header=False) -> list:
        cond = ts["cond"]
        if not cond:
            return []
        out = []
        first_row = look["firstRow"] and ri == 0
        last_row = look["lastRow"] and ri == nrows - 1
        first_col = look["firstColumn"] and col == 0
        last_col = look["lastColumn"] and col + span >= ncols
        if not look["noVBand"] and not first_col and not last_col:
            idx = col - (1 if look["firstColumn"] else 0)
            out.append(cond.get("band1Vert" if idx % 2 == 0 else "band2Vert"))
        if not look["noHBand"] and not first_row and not last_row:
            idx = ri - (1 if look["firstRow"] else 0)
            out.append(cond.get("band1Horz" if idx % 2 == 0 else "band2Horz"))
        if first_col:
            out.append(cond.get("firstCol"))
        if last_col:
            out.append(cond.get("lastCol"))
        if first_row:
            out.append(cond.get("firstRow"))
        if last_row:
            out.append(cond.get("lastRow"))
        if first_row and first_col:
            out.append(cond.get("nwCell"))
        if first_row and last_col:
            out.append(cond.get("neCell"))
        if last_row and first_col:
            out.append(cond.get("swCell"))
        if last_row and last_col:
            out.append(cond.get("seCell"))
        return [c for c in out if c]

    @staticmethod
    def _cell_borders(tb, cond_b, cell_b, ri, nrows, col, span, ncols) -> dict:
        def pick(*keys):
            for src in (cell_b, cond_b):
                for k in keys:
                    if k in src:
                        return src[k]
            return None

        out = {}
        for side, edge, inner, at_edge in (("top", "top", "insideH", ri == 0),
                                           ("bottom", "bottom", "insideH", ri == nrows - 1),
                                           ("left", "left", "insideV", col == 0),
                                           ("right", "right", "insideV", col + span >= ncols)):
            if side in cell_b:
                out[side] = cell_b[side]
            elif side in cond_b:
                out[side] = cond_b[side]
            else:
                key = edge if at_edge else inner
                if not at_edge and inner in cond_b:
                    out[side] = cond_b[inner]
                else:
                    out[side] = tb.get(key)
        return out


def _hyperlink_url(instr: str) -> str | None:
    """The external target of a HYPERLINK field, None for a jump inside the document (\\l)."""
    rest = instr.strip()[len("HYPERLINK"):].strip()
    if not rest or rest.startswith("\\l"):
        return None
    m = re.match(r'"([^"]*)"', rest) or re.match(r"(\S+)", rest)
    if not m:
        return None
    url = m.group(1)
    anchor = re.search(r'\\l\s+"([^"]*)"', rest)
    return url + ("#" + anchor.group(1) if anchor else "")


def _vml_size(style: str) -> tuple[int, int]:
    """width/height from a VML style string -> EMU."""
    def length(name):
        m = re.search(name + r"\s*:\s*([\d.]+)\s*(pt|in|cm|mm|px)?", style)
        if not m:
            return 0
        v = float(m.group(1))
        unit = m.group(2) or "px"
        per = {"pt": 12700, "in": 914400, "cm": 360000, "mm": 36000, "px": 9525}[unit]
        return int(v * per)
    return length("width"), length("height")
