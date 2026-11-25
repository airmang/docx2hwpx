"""Write the intermediate model as .hwpx with python-hwpx.

Public engine API is used wherever it exists; the few raw OWPML touches (page-number autoNum,
list head attributes, outline numbering reference, empty-run cleanup) copy the shape desktop
Hangul writes when it converts a .docx itself. No line-layout cache (hp:linesegarray) is written:
Hangul lays the text out when it opens the file.
"""
from __future__ import annotations

import hashlib
import re
from collections import Counter

from hwpx import HwpxDocument
from hwpx.objects import LineSpacing
from lxml import etree

from . import model as M
from .losses import LossLog

HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"
HH = "{http://www.hancom.co.kr/hwpml/2011/head}"

DEFAULT_FACE = "맑은 고딕"
LANGS_EA = ("HANGUL", "HANJA", "JAPANESE", "SYMBOL")
LANGS_LATIN = ("LATIN", "OTHER", "USER")

# What desktop Hangul writes into a document it converts from Word: Word-compatible layout (line
# height from the font as Word does, spacing after the last paragraph, table anchoring, ...).
WORD_LAYOUT_COMPAT = (
    "applyFontWeightToBold", "useInnerUnderline", "useLowercaseStrikeout", "extendLineheightToOffset",
    "treatQuotationAsLatin", "doNotAlignWhitespaceOnRight", "doNotAdjustWordInJustify", "baseCharUnitOnEAsian",
    "baseCharUnitOfIndentOnFirstChar", "adjustLineheightToFont", "adjustBaselineInFixedLinespacing",
    "applyPrevspacingBeneathObject", "applyNextspacingOfLastPara", "adjustParaBorderfillToSpacing",
    "connectParaBorderfillOfEqualBorder", "adjustParaBorderOffsetWithBorder", "extendLineheightToParaBorderOffset",
    "applyParaBorderToOutside", "applyMinColumnWidthTo1mm", "applyTabPosBasedOnSegment", "breakTabOverLine",
    "adjustVertPosOfLine", "doNotAlignLastForbidden", "adjustMarginFromAdjustLineheight",
    "baseLineSpacingOnLineGrid", "applyCharSpacingToCharGrid", "doNotApplyGridInHeaderFooter",
    "applyExtendHeaderFooterEachSection", "doNotApplyLinegridAtNoLinespacing", "doNotAdjustEmptyAnchorLine",
    "overlapBothAllowOverlap", "extendVertLimitToPageMargins", "doNotHoldAnchorOfTable",
    "doNotFormattingAtBeneathAnchor", "adjustBaselineOfObjectToBottom",
)

# English names of fonts that carry Hangul glyphs
KOREAN_FONTS_EN = {
    "malgun gothic", "gulim", "gulimche", "dotum", "dotumche", "batang", "batangche", "gungsuh", "gungsuhche",
    "nanumgothic", "nanummyeongjo", "nanumbarungothic", "nanumsquare", "noto sans kr", "noto serif kr",
    "noto sans cjk kr", "noto serif cjk kr", "pretendard", "spoqa han sans", "spoqa han sans neo",
    "source han sans k", "source han serif k", "apple sd gothic neo", "hcr batang", "hcr dotum",
    "hancom gothic", "kopubworld batang", "kopubworld dotum", "ms gothic", "ms mincho", "simsun",
    "microsoft yahei", "meiryo", "yu gothic", "yu mincho", "arial unicode ms",
}

HIGHLIGHT = {
    "yellow": "#FFFF00", "green": "#00FF00", "cyan": "#00FFFF", "magenta": "#FF00FF", "blue": "#0000FF",
    "red": "#FF0000", "darkBlue": "#000080", "darkCyan": "#008080", "darkGreen": "#008000",
    "darkMagenta": "#800080", "darkRed": "#800000", "darkYellow": "#808000", "darkGray": "#808080",
    "lightGray": "#C0C0C0", "black": "#000000", "white": "#FFFFFF",
}

UNDERLINE = {
    "single": "SOLID", "words": "SOLID", "double": "DOUBLE_SLIM", "thick": "SOLID", "dotted": "DOT",
    "dottedHeavy": "DOT", "dash": "DASH", "dashedHeavy": "DASH", "dashLong": "LONG_DASH",
    "dashLongHeavy": "LONG_DASH", "dotDash": "DASH_DOT", "dashDotHeavy": "DASH_DOT", "dotDotDash": "DASH_DOT_DOT",
    "dashDotDotHeavy": "DASH_DOT_DOT", "wave": "WAVE", "wavyHeavy": "WAVE", "wavyDouble": "DOUBLE_WAVE",
}

BORDER_TYPE = {
    "single": "SOLID", "thick": "SOLID", "double": "DOUBLE_SLIM", "dotted": "DOT", "dashed": "DASH",
    "dashSmallGap": "DASH", "dotDash": "DASH_DOT", "dotDotDash": "DASH_DOT_DOT", "triple": "SLIM_THICK_SLIM",
    "thinThickSmallGap": "SLIM_THICK", "thickThinSmallGap": "THICK_SLIM", "thinThickMediumGap": "SLIM_THICK",
    "thickThinMediumGap": "THICK_SLIM", "thinThickLargeGap": "SLIM_THICK", "thickThinLargeGap": "THICK_SLIM",
}
LINE_WIDTHS = (0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0)

ALIGN = {"left": "LEFT", "start": "LEFT", "center": "CENTER", "right": "RIGHT", "end": "RIGHT",
         "both": "JUSTIFY", "distribute": "DISTRIBUTE", "thaiDistribute": "DISTRIBUTE",
         "highKashida": "JUSTIFY", "mediumKashida": "JUSTIFY", "lowKashida": "JUSTIFY"}

NUM_FORMAT = {
    "decimal": "DIGIT", "decimalZero": "DIGIT", "decimalFullWidth": "DIGIT", "decimalHalfWidth": "DIGIT",
    "upperRoman": "ROMAN_CAPITAL", "lowerRoman": "ROMAN_SMALL", "upperLetter": "LATIN_CAPITAL",
    "lowerLetter": "LATIN_SMALL", "decimalEnclosedCircle": "CIRCLED_DIGIT",
    "decimalEnclosedCircleChinese": "CIRCLED_DIGIT", "ganada": "HANGUL_SYLLABLE", "chosung": "HANGUL_JAMO",
    "koreanDigital": "HANGUL_PHONETIC", "koreanCounting": "HANGUL_PHONETIC", "koreanLegal": "HANGUL_PHONETIC",
    "koreanDigital2": "IDEOGRAPH", "ideographDigital": "IDEOGRAPH", "chineseCounting": "IDEOGRAPH",
    "japaneseCounting": "IDEOGRAPH", "ideographTraditional": "IDEOGRAPH",
}
PAGE_FORMAT = {"lowerRoman": "ROMAN_SMALL", "upperRoman": "ROMAN_CAPITAL", "lowerLetter": "LATIN_SMALL",
               "upperLetter": "LATIN_CAPITAL", "ganada": "HANGUL_SYLLABLE", "decimalEnclosedCircle": "CIRCLED_DIGIT"}


def hu(twips) -> int:
    """twips -> HWPUNIT"""
    return int(round((twips or 0) * 5))


def emu_hu(emu) -> int:
    return max(1, int(round(emu / 127)))


def has_hangul_glyphs(face: str) -> bool:
    if any("가" <= c <= "힣" or "ㄱ" <= c <= "ㆎ" for c in face):
        return True
    low = face.lower().strip()
    return low in KOREAN_FONTS_EN or low.startswith(("hy", "hancom", "함초롬", "nanum", "noto sans kr",
                                                     "noto serif kr", "kopub", "hcr"))


def _snap_width(sz_eighths: int) -> float:
    mm = (sz_eighths or 4) / 8 * 0.3528
    best = LINE_WIDTHS[0]
    for w in LINE_WIDTHS:
        if w <= mm + 0.005:
            best = w
    return best


def _color(v, default="#000000") -> str:
    if not v or str(v).lower() == "auto" or not re.fullmatch(r"[0-9A-Fa-f]{6}", str(v)):
        return default
    return "#" + str(v).upper()


class Sink:
    """Where paragraphs go: a body section, a table cell, a header/footer or a note.

    The container's existing first paragraph (it holds the section settings in a body section) is
    reused for the first paragraph written. Tables wait for the next paragraph and are anchored at
    its start, the way Hangul itself anchors a converted Word table.
    """

    def __init__(self, writer: "HwpxWriter", make, first=None, keep_first_runs=False):
        self.w = writer
        self.make = make
        self.first = first
        self.keep_first_runs = keep_first_runs
        self.pending: list = []
        self.pending_columns = None
        self.count = 0

    def paragraph(self, pp: str, cp: str, *, page_break=False, column_break=False):
        if self.pending and (page_break or column_break):
            self.flush()
        p = self._next(pp, cp)
        if page_break:
            p.page_break = True
        if column_break:
            p.column_break = True
        if self.pending_columns is not None:
            sec = self.pending_columns
            self.pending_columns = None
            self.w.add_columns(p, sec)
        tables, self.pending = self.pending, []
        for t in tables:
            self.w.write_table(p, t, cp)
        return p

    def _next(self, pp, cp):
        self.count += 1
        if self.first is not None:
            p, self.first = self.first, None
            p.para_pr_id_ref = pp
            self.w.clean_runs(p, cp, keep_controls=True)
            return p
        return self.make(pp, cp)

    def flush(self):
        if self.pending:
            pp, cp = self.w.default_shapes()
            tables, self.pending = self.pending, []
            for t in tables:
                p = self._next(pp, cp)
                self.w.write_table(p, t, cp)
                self.w.ensure_text_run(p, cp)

    def close(self):
        self.flush()
        if self.first is not None:
            pp, cp = self.w.default_shapes()
            p = self._next(pp, cp)
            self.w.ensure_text_run(p, cp)


class HwpxWriter:
    def __init__(self, model: M.Document, losses: LossLog, progress=None):
        self.model = model
        self.losses = losses
        self.progress = progress
        self.doc = HwpxDocument.new()
        self.styles = self.doc.styles
        self._cp: dict = {}
        self._pp: dict = {}
        self._bf: dict = {}
        self._img: dict = {}
        self._num: dict = {}
        self._bullet: dict = {}
        self._outline_id = None
        self._section_pg_fmt = "DIGIT"
        self._default = None
        self._para_done = 0
        self._para_total = max(1, _count_paragraphs(model))
        self.header_el = self.doc.oxml.headers[0].element

    # ---------------------------------------------------------------- shapes
    def char_shape(self, rpr: dict) -> str:
        ea = rpr.get("font_eastAsia")
        latin = rpr.get("font_ascii") or rpr.get("font_hAnsi")
        if not ea or not has_hangul_glyphs(ea):
            ea = latin if latin and has_hangul_glyphs(latin) else DEFAULT_FACE
        if not latin:
            latin = ea
        size = (rpr.get("sz") or 20) / 2
        u = rpr.get("u")
        underline = bool(u) and u not in ("none", "0", "false")
        color = _color(rpr.get("color"))
        hl = rpr.get("highlight")
        shade = HIGHLIGHT.get(hl) if hl and hl != "none" else None
        if shade is None and rpr.get("shd"):
            shade = _color(rpr.get("shd"), None)
        va = rpr.get("vertAlign")
        script = "sup" if va == "superscript" else "sub" if va == "subscript" else None
        spacing = rpr.get("spacing") or 0
        letter = int(round(spacing / (size * 20) * 100)) if spacing else 0
        letter = max(-50, min(50, letter))
        ratio = max(50, min(200, rpr.get("w") or 100))
        key = (ea, latin, size, bool(rpr.get("b")), bool(rpr.get("i")), underline, u if underline else None,
               bool(rpr.get("strike") or rpr.get("dstrike")), color, shade, script, letter, ratio,
               bool(rpr.get("emboss")), bool(rpr.get("imprint")), bool(rpr.get("outline")),
               bool(rpr.get("shadow")))
        cid = self._cp.get(key)
        if cid is not None:
            return cid
        kwargs = dict(
            base_char_pr_id_ref=0, size=size, bold=key[3], italic=key[4], underline=underline, color=color,
            strike=key[7], ratio=ratio, letter_spacing=letter, emboss=key[13], engrave=key[14],
            font_by_lang={**{lang: ea for lang in LANGS_EA}, **{lang: latin for lang in LANGS_LATIN}},
        )
        if underline:
            kwargs["underline_shape"] = UNDERLINE.get(u, "SOLID")
        if shade:
            kwargs["highlight"] = shade
        if script:
            kwargs["script"] = script
        if key[15]:
            kwargs["outline"] = "SOLID"
        try:
            cid = self.styles.ensure_run(**kwargs)
        except Exception:  # noqa: BLE001 - an exotic value the engine refuses: keep the plain shape
            kwargs.pop("underline_shape", None)
            kwargs.pop("outline", None)
            cid = self.styles.ensure_run(**kwargs)
        cid = str(cid)
        self._cp[key] = cid
        return cid

    def para_shape(self, para: M.Paragraph, prev=None, nxt=None, *, heading=None) -> str:
        ppr = para.ppr
        left = ppr.get("left") or 0
        right = ppr.get("right") or 0
        first = ppr.get("first") or 0
        size_pt = (para.mark_rpr.get("sz") or 20) / 2
        if not left and ppr.get("leftChars"):
            left = int(ppr["leftChars"] / 100 * size_pt * 20)
        if not right and ppr.get("rightChars"):
            right = int(ppr["rightChars"] / 100 * size_pt * 20)
        if not first and ppr.get("hangingChars"):
            first = -int(ppr["hangingChars"] / 100 * size_pt * 20)
        elif not first and ppr.get("firstLineChars"):
            first = int(ppr["firstLineChars"] / 100 * size_pt * 20)
        if first < 0:
            margins = {"left": hu(left + first), "intent": hu(first)}
        else:
            margins = {"left": hu(left), "intent": hu(first)}
        margins["right"] = hu(right)
        before = ppr.get("before") or 0
        after = ppr.get("after") or 0
        if ppr.get("beforeLines") is not None and not ppr.get("before"):
            before = int(ppr["beforeLines"] * 240 / 100)
        if ppr.get("afterLines") is not None and not ppr.get("after"):
            after = int(ppr["afterLines"] * 240 / 100)
        if ppr.get("beforeAutospacing"):
            before = 280
        if ppr.get("afterAutospacing"):
            after = 280
        if ppr.get("contextualSpacing"):
            if isinstance(prev, M.Paragraph) and prev.style_id == para.style_id:
                before = 0
            if isinstance(nxt, M.Paragraph) and nxt.style_id == para.style_id:
                after = 0
        margins["prev"] = hu(before)
        margins["next"] = hu(after)
        rule = ppr.get("lineRule") or "auto"
        line = ppr.get("line")
        if line is None:
            ls = LineSpacing("PERCENT", 100)
        elif rule == "exact":
            ls = LineSpacing("FIXED", hu(abs(line)))
        elif rule == "atLeast":
            ls = LineSpacing("AT_LEAST", hu(line))
        else:
            ls = LineSpacing("PERCENT", max(10, int(round(line / 240 * 100))))
        align = ALIGN.get(ppr.get("jc") or "left", "LEFT")
        brk = {
            "keep_with_next": bool(ppr.get("keepNext")),
            "keep_lines": bool(ppr.get("keepLines")),
            "page_break_before": bool(ppr.get("pageBreakBefore")),
            "widow_orphan": bool(ppr.get("widowControl")),
        }
        if heading is None:
            heading = {"type": "NONE", "idRef": 0, "level": 0}
        tab_id = self.tab_definition(ppr.get("tabs"), left + min(first, 0))
        key = (align, tuple(sorted(margins.items())), ls.type.value, ls.value, tuple(sorted(brk.items())),
               tuple(sorted(heading.items())), tab_id)
        pid = self._pp.get(key)
        if pid is not None:
            return pid
        pid = str(self.styles.ensure_paragraph_format(
            base_para_pr_id_ref=0, alignment=align, margins=margins, heading=heading, break_setting=brk,
            line_spacing=ls, tab_pr_id_ref=tab_id))
        self._pp[key] = pid
        return pid

    def tab_definition(self, tabs, left_twips: int):
        if not tabs:
            return None
        stops = []
        for val, pos, leader in tabs:
            if val in ("clear", "bar", "num"):
                continue
            kind = {"left": "LEFT", "start": "LEFT", "right": "RIGHT", "end": "RIGHT", "center": "CENTER",
                    "decimal": "DECIMAL"}.get(val or "left", "LEFT")
            lead = {"dot": "DOT", "hyphen": "DASH", "underscore": "SOLID", "middleDot": "DOT"}.get(leader or "",
                                                                                               "NONE")
            # Word measures stops from the text margin, Hangul from the paragraph's left edge
            rel = pos - max(0, left_twips)
            if rel <= 0:
                continue
            stops.append({"pos_mm": round(rel / 56.6929, 2), "type": kind, "leader": lead})
        if not stops:
            return None
        key = tuple((s["pos_mm"], s["type"], s["leader"]) for s in stops)
        if key in self._pp:
            return self._pp[key]
        try:
            tid = str(self.styles.ensure_tab_definition(tab_stops=stops))
        except Exception:  # noqa: BLE001
            tid = None
        self._pp[key] = tid
        return tid

    def default_shapes(self):
        if self._default is None:
            ss = self.model_styles_default
            self._default = (self.para_shape(ss), self.char_shape(ss.mark_rpr))
        return self._default

    def border_fill(self, borders: dict, fill: str | None) -> str:
        sides = {}
        for side in ("left", "right", "top", "bottom"):
            b = borders.get(side)
            if not b or b[0] in ("nil", "none"):
                sides[side] = None
            else:
                sides[side] = {"type": BORDER_TYPE.get(b[0], "SOLID"), "width_mm": _snap_width(b[1]),
                               "color": _color(b[2])}
        fill_c = _color(fill, None) if fill else None
        key = (tuple((k, tuple(sorted(v.items())) if v else None) for k, v in sorted(sides.items())), fill_c)
        bid = self._bf.get(key)
        if bid is None:
            kwargs = {"sides": sides}
            if fill_c:
                kwargs["fill_color"] = fill_c
            bid = str(self.styles.ensure_border_fill(**kwargs))
            self._bf[key] = bid
        return bid

    # ---------------------------------------------------------------- lists
    def list_heading(self, para: M.Paragraph):
        nid, ilvl = para.num
        levels = self.model.numbering.get(nid) or {}
        lvl = levels.get(ilvl)
        if lvl is None:
            return None
        if lvl["fmt"] == "none":
            return None
        if lvl["fmt"] == "bullet":
            from .docx_reader import map_symbol_text
            ch = map_symbol_text(lvl["text"] or "•")[:1] or "•"
            if ch in ("o",) and (lvl.get("font") or "").lower().startswith("courier"):
                ch = "○"
            bid = self._bullet.get(ch)
            if bid is None:
                bid = str(self.styles.ensure_bullet(ch))
                self._patch_head(f"{HH}bullet", bid)
                self._bullet[ch] = bid
            return {"type": "BULLET", "idRef": int(bid), "level": 0}
        hid = self._num.get(nid)
        if hid is None:
            spec = []
            for i in range(10):
                lv = levels.get(i)
                if lv is None or i > 8:
                    spec.append({"text": ""})
                    continue
                text = re.sub(r"%(\d)", lambda m: "^" + m.group(1), lv["text"] or "")
                fmt = NUM_FORMAT.get(lv["fmt"], "DIGIT")
                spec.append({"text": text, "number_format": fmt, "start": max(0, min(65535, lv["start"]))})
            try:
                info = self.styles.add_numbering(spec)
            except Exception:  # noqa: BLE001 - a level pattern the engine refuses: plain digits
                info = self.styles.add_numbering([{"text": "^%d." % (i + 1)} for i in range(9)] + [{"text": ""}])
            hid = str(info.id)
            self._patch_head(f"{HH}numbering", hid)
            self._num[nid] = hid
        return {"type": "NUMBER", "idRef": int(hid), "level": min(ilvl, 9)}

    def _patch_head(self, tag: str, hid: str):
        """List heads as Hangul writes them for a converted Word list: no auto indent (the paragraph
        carries Word's hanging indent), width of the number itself, a 110% gap before the text."""
        for el in self.header_el.iter(tag):
            if el.get("id") == hid:
                for ph in el.iter(f"{HH}paraHead"):
                    ph.set("useInstWidth", "1")
                    ph.set("autoIndent", "0")
                    ph.set("textOffsetType", "PERCENT")
                    ph.set("textOffset", "110")
        self.doc.oxml.headers[0].mark_dirty()

    def outline_heading(self, level: int):
        if self._outline_id is None:
            info = self.styles.add_numbering([{"text": ""} for _ in range(10)])
            self._outline_id = str(info.id)
            for el in self.header_el.iter(f"{HH}numbering"):
                if el.get("id") == self._outline_id:
                    for ph in el.iter(f"{HH}paraHead"):
                        ph.set("autoIndent", "0")
                        ph.set("textOffset", "0")
            self.doc.oxml.headers[0].mark_dirty()
        return {"type": "OUTLINE", "idRef": 0, "level": max(0, min(level, 9))}

    # ---------------------------------------------------------------- document
    def write(self) -> HwpxDocument:
        secs = self.model.sections
        self.model_styles_default = M.Paragraph(ppr={}, mark_rpr={})
        if secs and secs[0].blocks:
            for b in secs[0].blocks:
                if isinstance(b, M.Paragraph):
                    self.model_styles_default = M.Paragraph(ppr={k: v for k, v in b.ppr.items()
                                                                 if k in ("line", "lineRule")},
                                                            mark_rpr=b.mark_rpr)
                    break
        prev_written = None
        hidx = -1
        cur_hdr = {"header": {}, "footer": {}}
        body_sink = None
        for i, sec in enumerate(secs):
            continuous = (i > 0 and sec.start_type == "continuous" and prev_written is not None
                          and _same_page(sec, prev_written))
            if not continuous:
                if body_sink is not None:
                    body_sink.close()
                if hidx >= 0:
                    self.doc.add_section()
                hidx += 1
                hsec = self.doc.sections[hidx]
                self.setup_page(hidx, sec)
                body_sink = Sink(self, lambda pp, cp, s=hsec: s.add_paragraph(
                    "", para_pr_id_ref=pp, char_pr_id_ref=cp, include_run=False, inherit_style=False),
                    first=hsec.paragraphs[0])
                if sec.cols > 1:
                    body_sink.pending_columns = sec
                self.write_headers(hidx, sec, cur_hdr)
            else:
                # a continuous break that only changes columns: a new column setting inside the section
                body_sink.flush()
                body_sink.pending_columns = sec
            prev_written = sec
            self.write_blocks(body_sink, sec.blocks)
        if body_sink is not None:
            body_sink.close()
        if self._outline_id is not None:
            for s in self.doc.sections:
                s.properties.element.set("outlineShapeIDRef", self._outline_id)
        self.set_word_compatibility()
        for s in self.doc.sections:
            s.mark_dirty()
        return self.doc

    def set_word_compatibility(self):
        """Mark the document as converted from Word, as Hangul does (no API for hh:compatibleDocument)."""
        head = self.header_el
        comp = head.find(f"{HH}compatibleDocument")
        if comp is None:
            comp = etree.Element(f"{HH}compatibleDocument")
            anchor = head.find(f"{HH}docOption")
            if anchor is not None:
                anchor.addprevious(comp)
            else:
                head.append(comp)
        comp.set("targetProgram", "MS_WORD")
        for child in list(comp):
            comp.remove(child)
        layout = etree.SubElement(comp, f"{HH}layoutCompatibility")
        for name in WORD_LAYOUT_COMPAT:
            etree.SubElement(layout, f"{HH}{name}")
        self.doc.oxml.headers[0].mark_dirty()

    def setup_page(self, idx: int, sec: M.Section):
        short, long_ = sorted((sec.page_w, sec.page_h))
        self.doc.page.set_size(width=hu(short), height=hu(long_),
                               orientation="landscape" if sec.landscape else "portrait", section=idx)
        m = sec.margins
        top, bottom = abs(m.get("top", 1440)), abs(m.get("bottom", 1440))
        hdr, ftr = m.get("header", 720), m.get("footer", 720)
        # Hangul's top margin runs to the header; the header area runs to the body
        h_top, h_header = (hdr, top - hdr) if top >= hdr else (top, 0)
        h_bottom, h_footer = (ftr, bottom - ftr) if bottom >= ftr else (bottom, 0)
        self.doc.page.set_margins(left=hu(m.get("left", 1440)), right=hu(m.get("right", 1440)), top=hu(h_top),
                                  bottom=hu(h_bottom), header=hu(h_header), footer=hu(h_footer),
                                  gutter=hu(m.get("gutter", 0)), section=idx)
        if sec.pg_start is not None:
            self.doc.page.set_start_numbers(page=sec.pg_start, section=idx)
        self._section_pg_fmt = PAGE_FORMAT.get(sec.pg_fmt or "", "DIGIT")

    def add_columns(self, p, sec: M.Section):
        kwargs = {"col_count": sec.cols, "same_gap": hu(sec.col_space)}
        if sec.col_sep:
            kwargs["separator_type"] = "SOLID"
        if sec.col_widths:
            kwargs["same_size"] = False
            kwargs["column_widths"] = [(hu(w), hu(s)) for w, s in sec.col_widths]
        try:
            p.add_column_definition(**kwargs)
        except Exception:  # noqa: BLE001
            kwargs.pop("column_widths", None)
            kwargs.pop("same_size", None)
            p.add_column_definition(**kwargs)

    def write_headers(self, idx: int, sec: M.Section, cur: dict):
        for kind, hfs in (("header", sec.headers), ("footer", sec.footers)):
            wanted = {}
            default, first, even = hfs.default, hfs.first, hfs.even
            body = default if default else (first if sec.title_pg and first else None)
            if sec.title_pg and first and default and first is not default and _has_text(first):
                self.losses.add("first_page_header")
            if self.model.even_odd_headers and even is not None:
                wanted["EVEN"] = even
                wanted["ODD"] = body
            else:
                wanted["BOTH"] = body
            hide_first = bool(sec.title_pg and not _has_text(first or []) and _has_text(body or []))
            vis = {"hide_first_header" if kind == "header" else "hide_first_footer": hide_first}
            self.doc.page.set_visibility(section=idx, **vis)
            for ptype, blocks in wanted.items():
                if cur[kind].get(ptype) is blocks:
                    continue
                if blocks is None and cur[kind].get(ptype) is None:
                    continue
                setter = self.doc.page.set_header if kind == "header" else self.doc.page.set_footer
                hf = setter(text="", page_type=ptype, section=idx)
                sink = Sink(self, lambda pp, cp, h=hf: self._hf_paragraph(h, pp, cp), first=hf.paragraphs[0])
                self.write_blocks(sink, blocks or [])
                sink.close()
                cur[kind][ptype] = blocks

    def _hf_paragraph(self, hf, pp, cp):
        p = hf.add_paragraph("", para_pr_id_ref=pp, char_pr_id_ref=cp)
        self.clean_runs(p, cp)
        return p

    # ---------------------------------------------------------------- blocks
    def write_blocks(self, sink: Sink, blocks: list):
        n = len(blocks)
        for i, b in enumerate(blocks):
            if isinstance(b, M.Paragraph):
                prev = blocks[i - 1] if i > 0 else None
                nxt = blocks[i + 1] if i + 1 < n else None
                self.write_paragraph(sink, b, prev, nxt)
                if b.after:
                    self.write_blocks(sink, b.after)
            elif isinstance(b, M.Table):
                if sink.pending:
                    sink.flush()
                sink.pending.append(b)

    def write_paragraph(self, sink: Sink, para: M.Paragraph, prev, nxt):
        heading = None
        if para.num:
            heading = self.list_heading(para)
        if heading is None:
            lvl = para.ppr.get("outlineLvl")
            if lvl is not None and 0 <= lvl < 9:
                heading = self.outline_heading(lvl)
        pp = self.para_shape(para, prev, nxt, heading=heading)
        cp = self.char_shape(para.mark_rpr)
        p = sink.paragraph(pp, cp, page_break=para.page_break, column_break=para.column_break)
        self.write_items(p, para.items)
        self.ensure_text_run(p, cp)
        self._para_done += 1
        if self.progress and self._para_done % 200 == 0:
            self.progress(0.3 + 0.55 * min(1.0, self._para_done / self._para_total), "글을 옮기는 중")

    def write_items(self, p, items: list):
        buf: list[str] = []
        buf_cp = None
        for it in items + [None]:
            if isinstance(it, M.Text):
                cp = self.char_shape(it.rpr)
                if cp == buf_cp:
                    buf.append(it.text)
                    continue
            if buf:
                p.add_run("".join(buf), char_pr_id_ref=buf_cp, expand_special_characters=True)
                buf, buf_cp = [], None
            if it is None:
                break
            if isinstance(it, M.Text):
                buf, buf_cp = [it.text], self.char_shape(it.rpr)
            elif isinstance(it, M.Field):
                self.add_auto_number(p, it)
            elif isinstance(it, M.NoteRef):
                self.add_note(p, it)
            elif isinstance(it, M.Hyperlink):
                self.add_hyperlink(p, it)
            elif isinstance(it, M.Image):
                self.add_image(p, it)

    def add_auto_number(self, p, f: M.Field):
        """Page number as Hangul writes it for a converted PAGE field: an autoNum control in a run."""
        cp = self.char_shape(f.rpr)
        run = etree.SubElement(p.element, f"{HP}run", {"charPrIDRef": cp})
        ctrl = etree.SubElement(run, f"{HP}ctrl")
        an = etree.SubElement(ctrl, f"{HP}autoNum", {"num": "1",
                                                     "numType": "PAGE" if f.kind == "PAGE" else "TOTAL_PAGE"})
        fmt = f.fmt if f.fmt != "DIGIT" else self._section_pg_fmt
        etree.SubElement(an, f"{HP}autoNumFormat", {"type": fmt, "userChar": "", "prefixChar": "",
                                                    "suffixChar": "", "supscript": "0"})

    def add_note(self, p, ref: M.NoteRef):
        cp = self.char_shape(ref.rpr)
        if ref.kind == "footnote":
            note = p.add_footnote("", char_pr_id_ref=cp)
        else:
            note = p.add_endnote("", char_pr_id_ref=cp)
        first = note.paragraphs[0] if note.paragraphs else None

        def make(pp, cp_, n=note):
            q = n.add_paragraph("", para_pr_id_ref=pp, char_pr_id_ref=cp_)
            self.clean_runs(q, cp_)
            return q

        sink = Sink(self, make, first=first)
        self.write_blocks(sink, ref.blocks)
        sink.close()

    def add_hyperlink(self, p, link: M.Hyperlink):
        texts = [it for it in link.items if isinstance(it, M.Text)]
        others = [it for it in link.items if not isinstance(it, M.Text)]
        if link.url and texts:
            text = "".join(t.text for t in texts).replace("\n", " ").replace("\t", " ")
            cp = self.char_shape(texts[0].rpr)
            try:
                p.add_hyperlink(link.url, text, char_pr_id_ref=cp)
                # the field's begin/end runs take the link's character shape, as in Hangul's own files
                for run in p.element.findall(f"{HP}run"):
                    if run.get("charPrIDRef") == "0" and (run.find(f"{HP}ctrl/{HP}fieldBegin") is not None
                                                         or run.find(f"{HP}ctrl/{HP}fieldEnd") is not None):
                        run.set("charPrIDRef", cp)
            except Exception:  # noqa: BLE001 - a target the engine refuses: keep the words
                self.write_items(p, texts)
            if others:
                self.write_items(p, others)
        else:
            self.write_items(p, link.items)

    def add_image(self, p, img: M.Image):
        digest = hashlib.sha1(img.data).hexdigest()
        item = self._img.get(digest)
        if item is None:
            try:
                item = self.doc.media.add_image(img.data, img.fmt)
            except Exception:  # noqa: BLE001 - bytes that are not the format the name says
                self.losses.add("image_unsupported")
                return
            self._img[digest] = item
        p.add_picture(item, width=emu_hu(img.cx), height=emu_hu(img.cy), treat_as_char=True,
                      char_pr_id_ref=self.char_shape(img.rpr))

    # ---------------------------------------------------------------- tables
    def write_table(self, p, t: M.Table, cp: str):
        ncols = max([len(t.grid)] + [sum(c.span for c in r.cells) for r in t.rows]) if t.rows else 0
        nrows = len(t.rows)
        if not nrows or not ncols:
            return
        widths = _column_widths(t, ncols)
        nob = self.border_fill({}, None)
        tbl = p.add_table(nrows, ncols, width=hu(sum(widths)), border_fill_id_ref=nob, char_pr_id_ref=cp)
        pos = {"treat_as_char": False, "text_wrap": "TOP_AND_BOTTOM", "flow_with_text": True,
               "allow_overlap": False, "horizontal_relative_to": "COLUMN", "vertical_relative_to": "PARA"}
        if t.jc in ("center", "right", "end"):
            pos["horizontal_alignment"] = "CENTER" if t.jc == "center" else "RIGHT"
        elif t.indent:
            pos["horizontal_offset"] = hu(t.indent)
        tbl.set_position(**pos)
        cm = t.cell_margins
        tbl.set_in_margins(left=hu(cm.get("left", 108)), right=hu(cm.get("right", 108)),
                           top=hu(cm.get("top", 0)), bottom=hu(cm.get("bottom", 0)))
        tbl.set_column_widths([hu(w) for w in widths])
        # place Word cells on the grid
        origin = {}
        for r, row in enumerate(t.rows):
            c = 0
            for cell in row.cells:
                if cell.vmerge == "continue":
                    for rr in range(r - 1, -1, -1):
                        o = origin.get((rr, c))
                        if o is not None:
                            o[2] = r - o[0] + 1
                            break
                else:
                    origin[(r, c)] = [r, c, 1, cell]
                c += cell.span
        merges = []
        for (r, c), (r0, c0, rs, cell) in origin.items():
            if rs > 1 or cell.span > 1:
                merges.append((r, c, r + rs - 1, min(ncols - 1, c + cell.span - 1)))
        for m in merges:
            tbl.merge_cells(*m)
        header_rows = 0
        for r, row in enumerate(t.rows):
            if row.header and r == header_rows:
                header_rows += 1
        covered = set()
        for (r, c), (r0, c0, rs, cell) in origin.items():
            for rr in range(r, r + rs):
                for cc in range(c, min(ncols, c + cell.span)):
                    covered.add((rr, cc))
        for r in range(nrows):
            for c in range(ncols):
                if (r, c) not in covered:
                    tbl.set_cell_border_fill(r, c, nob)
                    hc = tbl.cell(r, c)
                    hc.vertical_alignment = "TOP"
        for (r, c), (r0, c0, rs, cell) in sorted(origin.items()):
            hc = tbl.cell(r, c)
            tbl.set_cell_border_fill(r, c, self.border_fill(cell.borders, cell.fill))
            hc.vertical_alignment = {"center": "CENTER", "bottom": "BOTTOM"}.get(cell.valign or "top", "TOP")
            if cell.margins:
                hc.set_margins(**{k: hu(v) for k, v in cell.margins.items() if k in ("left", "right", "top",
                                                                                    "bottom")})
            row = t.rows[r]
            if row.height and rs == 1:
                hc.set_size(height=hu(row.height))
            first = hc.paragraphs[0] if hc.paragraphs else None

            def make(pp, cp, cell_=hc):
                q = cell_.add_paragraph("", para_pr_id_ref=pp, char_pr_id_ref=cp)
                self.clean_runs(q, cp)
                return q

            sink = Sink(self, make, first=first)
            self.write_blocks(sink, cell.blocks)
            sink.close()
        if header_rows:
            try:
                tbl.set_header_rows(header_rows, repeat=True)
            except Exception:  # noqa: BLE001
                pass

    # ---------------------------------------------------------------- raw helpers
    @staticmethod
    def clean_runs(p, cp, keep_controls=False):
        """Drop the empty text runs the engine seeds a new paragraph with; give the run that holds
        section/column controls the paragraph's character shape, as Hangul does."""
        el = p.element
        for run in list(el.findall(f"{HP}run")):
            for t in list(run.findall(f"{HP}t")):
                if not (t.text or "") and len(t) == 0:
                    run.remove(t)
            if len(run) == 0:
                el.remove(run)
            elif keep_controls and run.find(f"{HP}secPr") is not None:
                run.set("charPrIDRef", cp)
            elif run.find(f"{HP}ctrl/{HP}autoNum") is not None and len(run) == 1:
                run.set("charPrIDRef", cp)

    @staticmethod
    def ensure_text_run(p, cp):
        el = p.element
        for run in el.findall(f"{HP}run"):
            if run.find(f"{HP}t") is not None or run.find(f"{HP}pic") is not None \
                    or run.find(f"{HP}tbl") is not None or run.find(f"{HP}ctrl") is not None:
                return
        run = etree.SubElement(el, f"{HP}run", {"charPrIDRef": cp})
        etree.SubElement(run, f"{HP}t")


def _has_text(blocks) -> bool:
    for b in blocks or []:
        if isinstance(b, M.Paragraph):
            if b.items or b.after:
                return True
        elif isinstance(b, M.Table):
            return True
    return False


def _same_page(a: M.Section, b: M.Section) -> bool:
    return (a.page_w, a.page_h, a.landscape, tuple(sorted(a.margins.items()))) == \
        (b.page_w, b.page_h, b.landscape, tuple(sorted(b.margins.items())))


def _column_widths(t: M.Table, ncols: int) -> list[int]:
    """Column widths as Word lays them out: the cells' preferred widths (tcW) where single cells give
    them, the grid otherwise."""
    grid = list(t.grid) + [0] * (ncols - len(t.grid))
    seen: dict[int, Counter] = {}
    for row in t.rows:
        c = 0
        for cell in row.cells:
            if cell.span == 1 and cell.vmerge != "continue" and cell.width:
                seen.setdefault(c, Counter())[cell.width] += 1
            c += cell.span
    widths = []
    for c in range(ncols):
        if c in seen:
            widths.append(seen[c].most_common(1)[0][0])
        else:
            widths.append(grid[c] or 0)
    known = [w for w in widths if w]
    fill = int(sum(known) / len(known)) if known else 1700
    widths = [w or fill for w in widths]
    if t.width and t.width[1] == "dxa" and t.width[0] > 0 and not seen:
        total = sum(widths)
        widths = [int(w * t.width[0] / total) for w in widths]
    return [max(w, 100) for w in widths]


def _count_paragraphs(model: M.Document) -> int:
    n = 0
    stack = [b for s in model.sections for b in s.blocks]
    while stack:
        b = stack.pop()
        if isinstance(b, M.Paragraph):
            n += 1
        elif isinstance(b, M.Table):
            for row in b.rows:
                for c in row.cells:
                    stack.extend(c.blocks)
    return n
