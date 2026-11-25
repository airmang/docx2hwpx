"""Word style resolution: docDefaults -> style chain (basedOn) -> numbering -> direct formatting.

Paragraph properties (ppr) and run properties (rpr) are plain dicts holding only the keys a level
sets; `merge` lays a later level over an earlier one.
"""
from __future__ import annotations

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
W = "{%s}" % W_NS
A = "{%s}" % A_NS

TOGGLES = ("b", "i", "caps", "smallCaps", "strike", "dstrike", "vanish", "emboss", "imprint", "outline", "shadow")


def wval(el, attr="val", default=None):
    if el is None:
        return default
    v = el.get(W + attr)
    return default if v is None else v


def on(el) -> bool | None:
    """A Word on/off property: present without val (or val true/1/on) is on."""
    if el is None:
        return None
    v = el.get(W + "val")
    return v is None or v.lower() in ("1", "true", "on")


def to_int(v, default=None):
    if v is None:
        return default
    try:
        return int(round(float(v)))
    except ValueError:
        return default


class Theme:
    def __init__(self, xml: bytes | None):
        self.fonts = {"major": {}, "minor": {}}
        if not xml:
            return
        root = etree.fromstring(xml)
        for kind in ("major", "minor"):
            el = root.find(f".//{A}{kind}Font")
            if el is None:
                continue
            d = self.fonts[kind]
            for tag in ("latin", "ea", "cs"):
                t = el.find(A + tag)
                if t is not None and t.get("typeface"):
                    d[tag] = t.get("typeface")
            for f in el.findall(A + "font"):
                d["script:" + (f.get("script") or "")] = f.get("typeface")

    def font(self, theme_name: str) -> str | None:
        kind = "major" if theme_name.startswith("major") else "minor"
        d = self.fonts[kind]
        rest = theme_name[5:]
        if rest in ("HAnsi", "Ascii"):
            return d.get("latin")
        if rest == "EastAsia":
            return d.get("ea") or d.get("script:Hang") or d.get("script:Jpan") or d.get("script:Hans")
        if rest == "Bidi":
            return d.get("cs")
        return None


def parse_rpr(el, theme: Theme) -> dict:
    out: dict = {}
    if el is None:
        return out
    for child in el:
        tag = etree.QName(child).localname if isinstance(child.tag, str) else ""
        if tag == "rFonts":
            for slot in ("ascii", "hAnsi", "eastAsia", "cs"):
                th = child.get(W + slot + "Theme")
                name = theme.font(th) if th else None
                if name is None:
                    name = child.get(W + slot)
                if name:
                    out["font_" + slot] = name
        elif tag in TOGGLES:
            out[tag] = on(child)
        elif tag == "sz":
            v = to_int(wval(child))
            if v:
                out["sz"] = v
        elif tag == "u":
            out["u"] = wval(child, default="single")
        elif tag == "color":
            v = wval(child)
            if v:
                out["color"] = v
        elif tag == "highlight":
            out["highlight"] = wval(child)
        elif tag == "shd":
            fill = child.get(W + "fill")
            if fill and fill.lower() != "auto":
                out["shd"] = fill
            elif fill:
                out["shd"] = None
        elif tag == "vertAlign":
            out["vertAlign"] = wval(child)
        elif tag == "spacing":
            out["spacing"] = to_int(wval(child), 0)
        elif tag == "w":
            out["w"] = to_int(wval(child), 100)
        elif tag == "rStyle":
            out["rStyle"] = wval(child)
    return out


def parse_ppr(el) -> dict:
    out: dict = {}
    if el is None:
        return out
    for child in el:
        if not isinstance(child.tag, str):
            continue
        tag = etree.QName(child).localname
        if tag == "jc":
            out["jc"] = wval(child)
        elif tag == "ind":
            left = child.get(W + "left") or child.get(W + "start")
            right = child.get(W + "right") or child.get(W + "end")
            if left is not None:
                out["left"] = to_int(left, 0)
            if right is not None:
                out["right"] = to_int(right, 0)
            if child.get(W + "hanging") is not None:
                out["first"] = -to_int(child.get(W + "hanging"), 0)
            elif child.get(W + "firstLine") is not None:
                out["first"] = to_int(child.get(W + "firstLine"), 0)
            for k in ("leftChars", "startChars", "rightChars", "endChars", "hangingChars", "firstLineChars"):
                v = child.get(W + k)
                if v is not None and to_int(v, 0) != 0:
                    out[k.replace("start", "left").replace("end", "right")] = to_int(v, 0)
        elif tag == "spacing":
            for k in ("before", "after", "line", "beforeLines", "afterLines"):
                v = child.get(W + k)
                if v is not None:
                    out[k] = to_int(v, 0)
            if child.get(W + "lineRule") is not None:
                out["lineRule"] = child.get(W + "lineRule")
            for k in ("beforeAutospacing", "afterAutospacing"):
                v = child.get(W + k)
                if v is not None:
                    out[k] = v.lower() in ("1", "true", "on")
        elif tag in ("keepNext", "keepLines", "pageBreakBefore", "widowControl", "contextualSpacing"):
            out[tag] = on(child)
        elif tag == "outlineLvl":
            out["outlineLvl"] = to_int(wval(child), 9)
        elif tag == "numPr":
            nid = child.find(W + "numId")
            lvl = child.find(W + "ilvl")
            if nid is not None:
                out["numId"] = wval(nid)
            if lvl is not None:
                out["ilvl"] = to_int(wval(lvl), 0)
        elif tag == "pStyle":
            out["pStyle"] = wval(child)
        elif tag == "shd":
            fill = child.get(W + "fill")
            if fill and fill.lower() != "auto":
                out["shd"] = fill
        elif tag == "tabs":
            tabs = []
            for t in child.findall(W + "tab"):
                tabs.append((wval(t), to_int(t.get(W + "pos"), 0), t.get(W + "leader")))
            out["tabs"] = tabs
        elif tag == "textDirection":
            out["textDirection"] = wval(child)
    return out


def merge(base: dict, over: dict) -> dict:
    if not over:
        return base
    out = dict(base)
    if "tabs" in over and "tabs" in base:
        # tab lists add up; a "clear" tab removes an inherited stop at that position
        stops = {pos: (val, pos, leader) for val, pos, leader in base["tabs"]}
        for val, pos, leader in over["tabs"]:
            if val == "clear":
                stops.pop(pos, None)
            else:
                stops[pos] = (val, pos, leader)
        over = dict(over)
        over["tabs"] = sorted(stops.values(), key=lambda t: t[1])
    out.update(over)
    return out


def parse_borders(el) -> dict:
    """tblBorders / tcBorders -> side -> (val, sz eighths-pt, color) ; a side set to none maps to None."""
    out: dict = {}
    if el is None:
        return out
    for child in el:
        if not isinstance(child.tag, str):
            continue
        side = etree.QName(child).localname
        side = {"start": "left", "end": "right"}.get(side, side)
        val = child.get(W + "val") or "nil"
        if val in ("nil", "none"):
            out[side] = None
        else:
            out[side] = (val, to_int(child.get(W + "sz"), 4), child.get(W + "color") or "auto")
    return out


def parse_margins(el) -> dict:
    out: dict = {}
    if el is None:
        return out
    for child in el:
        if not isinstance(child.tag, str):
            continue
        side = etree.QName(child).localname
        side = {"start": "left", "end": "right"}.get(side, side)
        if child.get(W + "type") in (None, "dxa"):
            out[side] = to_int(child.get(W + "w"), 0)
    return out


class StyleSheet:
    def __init__(self, styles_xml: bytes | None, theme: Theme):
        self.theme = theme
        self.styles: dict[str, etree._Element] = {}
        self.doc_ppr: dict = {}
        self.doc_rpr: dict = {}
        self.default_para: str | None = None
        self.default_table: str | None = None
        self.default_char: str | None = None
        self._pcache: dict = {}
        self._ccache: dict = {}
        self._tcache: dict = {}
        if not styles_xml:
            return
        root = etree.fromstring(styles_xml)
        dd = root.find(W + "docDefaults")
        if dd is not None:
            self.doc_rpr = parse_rpr(dd.find(f"{W}rPrDefault/{W}rPr"), theme)
            self.doc_ppr = parse_ppr(dd.find(f"{W}pPrDefault/{W}pPr"))
        for st in root.findall(W + "style"):
            sid = st.get(W + "styleId")
            if not sid:
                continue
            self.styles[sid] = st
            if st.get(W + "default") in ("1", "true", "on"):
                kind = st.get(W + "type")
                if kind == "paragraph":
                    self.default_para = sid
                elif kind == "table":
                    self.default_table = sid
                elif kind == "character":
                    self.default_char = sid

    def name(self, sid: str | None) -> str:
        st = self.styles.get(sid or "")
        if st is None:
            return ""
        n = st.find(W + "name")
        return (wval(n) or "").lower()

    def _chain(self, sid: str | None, seen=None) -> list:
        out = []
        seen = seen or set()
        while sid and sid in self.styles and sid not in seen:
            seen.add(sid)
            out.append(self.styles[sid])
            sid = wval(self.styles[sid].find(W + "basedOn"))
        return list(reversed(out))

    def para_style(self, sid: str | None) -> tuple[dict, dict]:
        """(ppr, rpr) of a paragraph style, inherited along basedOn (without docDefaults)."""
        if sid not in self.styles:
            sid = self.default_para
        if sid in self._pcache:
            return self._pcache[sid]
        ppr: dict = {}
        rpr: dict = {}
        for st in self._chain(sid):
            p = parse_ppr(st.find(W + "pPr"))
            p.pop("pStyle", None)
            ppr = merge(ppr, p)
            rpr = merge(rpr, parse_rpr(st.find(W + "rPr"), self.theme))
        self._pcache[sid] = (ppr, rpr)
        return ppr, rpr

    def char_style(self, sid: str | None) -> dict:
        if not sid or sid not in self.styles:
            return {}
        if sid in self._ccache:
            return self._ccache[sid]
        rpr: dict = {}
        for st in self._chain(sid):
            rpr = merge(rpr, parse_rpr(st.find(W + "rPr"), self.theme))
        rpr.pop("rStyle", None)
        self._ccache[sid] = rpr
        return rpr

    def table_style(self, sid: str | None) -> dict:
        """Merged table style: ppr, rpr, borders, margins, shd and conditional parts."""
        if sid not in self.styles:
            sid = self.default_table
        if sid in self._tcache:
            return self._tcache[sid]
        ts = {"ppr": {}, "rpr": {}, "borders": {}, "margins": {}, "shd": None, "cond": {}}
        for st in self._chain(sid):
            ts["ppr"] = merge(ts["ppr"], parse_ppr(st.find(W + "pPr")))
            ts["rpr"] = merge(ts["rpr"], parse_rpr(st.find(W + "rPr"), self.theme))
            tblpr = st.find(W + "tblPr")
            if tblpr is not None:
                ts["borders"].update(parse_borders(tblpr.find(W + "tblBorders")))
                ts["margins"].update(parse_margins(tblpr.find(W + "tblCellMar")))
            tcpr = st.find(W + "tcPr")
            if tcpr is not None:
                shd = tcpr.find(W + "shd")
                if shd is not None and (shd.get(W + "fill") or "auto").lower() != "auto":
                    ts["shd"] = shd.get(W + "fill")
            for cond in st.findall(W + "tblStylePr"):
                kind = cond.get(W + "type")
                c = ts["cond"].setdefault(kind, {"ppr": {}, "rpr": {}, "borders": {}, "shd": None})
                c["ppr"] = merge(c["ppr"], parse_ppr(cond.find(W + "pPr")))
                c["rpr"] = merge(c["rpr"], parse_rpr(cond.find(W + "rPr"), self.theme))
                ctc = cond.find(W + "tcPr")
                if ctc is not None:
                    c["borders"].update(parse_borders(ctc.find(W + "tcBorders")))
                    shd = ctc.find(W + "shd")
                    if shd is not None and (shd.get(W + "fill") or "auto").lower() != "auto":
                        c["shd"] = shd.get(W + "fill")
        self._tcache[sid] = ts
        return ts
