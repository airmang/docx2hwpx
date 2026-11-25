# -*- coding: utf-8 -*-
"""Converter core: every fixture converts into a package desktop Hangul accepts, with the Word features
carried as Hangul's own conversion carries them; bad inputs fail with a Korean one-liner."""
from __future__ import annotations

import io
import re
import zipfile
from collections import Counter
from pathlib import Path

import hwpx
import pytest
from lxml import etree

from docx2hwpx import convert

FIX = Path(__file__).parent / "fixtures"
FIXTURES = sorted(p.stem for p in FIX.glob("*.docx"))
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def L(tag) -> str:
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def features(path: Path) -> dict:
    """Feature counts of an .hwpx (the same notions as the bench's analyze.py)."""
    z = zipfile.ZipFile(path)
    secs = sorted(n for n in z.namelist() if re.match(r"Contents/section\d+\.xml", n))
    els = [e for n in secs for e in etree.fromstring(z.read(n)).iter() if isinstance(e.tag, str)]
    head = etree.fromstring(z.read("Contents/header.xml"))
    by = Counter(L(e.tag) for e in els)
    pp = {e.get("id"): e for e in head.iter() if L(e.tag) == "paraPr"}
    cp = {e.get("id"): e for e in head.iter() if L(e.tag) == "charPr"}
    bf = {e.get("id"): e for e in head.iter() if L(e.tag) == "borderFill"}
    heads = Counter()
    for e in els:
        if L(e.tag) == "p":
            h = next(k for k in pp[e.get("paraPrIDRef")].iter() if L(k.tag) == "heading")
            heads[h.get("type")] += 1
    chars = Counter()
    for cid in {e.get("charPrIDRef") for e in els if L(e.tag) == "run" and e.findtext(f"{{{e.nsmap.get('hp')}}}t")}:
        kids = {L(k.tag): k for k in cp[cid]}
        chars["bold"] += "bold" in kids
        chars["italic"] += "italic" in kids
        chars["underline"] += kids["underline"].get("type") != "NONE"
        chars["strike"] += kids["strikeout"].get("shape") != "NONE"
        chars["sup"] += "supscript" in kids
        chars["sub"] += "subscript" in kids
        chars["color"] += cp[cid].get("textColor") not in ("#000000", None)
        chars["shade"] += cp[cid].get("shadeColor") not in ("none", None)
    fills = sum(1 for e in els if L(e.tag) == "tc" and any(
        L(k.tag) == "winBrush" and k.get("faceColor") not in ("none", None) for k in bf[e.get("borderFillIDRef")].iter()))
    page_prs = [e for e in els if L(e.tag) == "pagePr"]
    text = "".join(e.text or "" for e in els if L(e.tag) == "t")
    return {
        "tables": by["tbl"],
        "nested_table": sum(1 for e in els if L(e.tag) == "tbl" and any(L(a.tag) == "tc" for a in e.iterancestors())),
        "merged": sum(1 for e in els if L(e.tag) == "cellSpan"
                      and (e.get("colSpan") != "1" or e.get("rowSpan") != "1")),
        "pictures": by["pic"], "header": by["header"], "footer": by["footer"],
        "page_num": sum(1 for e in els if L(e.tag) == "autoNum" and e.get("numType") == "PAGE"),
        "footnotes": by["footNote"],
        "hyperlinks": sum(1 for e in els if L(e.tag) == "fieldBegin" and e.get("type") == "HYPERLINK"),
        "sections": len(page_prs),
        "landscape": sum(1 for e in page_prs if e.get("landscape") == "NARROWLY"),
        "two_columns": sum(1 for e in els if L(e.tag) == "colPr" and int(e.get("colCount", "1")) >= 2),
        "page_breaks": sum(1 for e in els if L(e.tag) == "p" and e.get("pageBreak") == "1"),
        "linesegarray": by["linesegarray"], "bullets": heads["BULLET"], "numbers": heads["NUMBER"],
        "outline": heads["OUTLINE"], "cell_fills": fills, "chars": chars, "text": text,
    }


def docx_text(path: Path) -> str:
    z = zipfile.ZipFile(path)
    out = []
    for n in sorted(z.namelist()):
        if re.match(r"word/(document|header\d*|footer\d*|footnotes)\.xml$", n):
            out.append("".join(t.text or "" for t in etree.fromstring(z.read(n)).iter(f"{{{W_NS}}}t")
                               # a page field's cached number is drawn live by Hangul
                               if L(t.getparent().getparent().tag) != "fldSimple"))
    return "".join(out)


@pytest.fixture(scope="module")
def converted(tmp_path_factory):
    out = tmp_path_factory.mktemp("hwpx")
    results = {}
    for stem in FIXTURES:
        steps = []
        r = convert(FIX / f"{stem}.docx", out / f"{stem}.hwpx", progress=lambda f, t: steps.append((f, t)))
        results[stem] = (r, steps)
    return results


@pytest.mark.parametrize("stem", FIXTURES)
def test_fixture_converts_and_is_safe(converted, stem):
    r, steps = converted[stem]
    assert r.ok, r.error
    assert r.output.exists() and r.seconds < 10
    assert steps and steps[0][0] == 0.0 and steps[-1][0] == 1.0
    pkg = hwpx.validate_package(str(r.output))
    assert pkg.ok, pkg
    safety = hwpx.validate_editor_open_safety(str(r.output))
    assert safety.ok, safety
    f = features(r.output)
    assert f["linesegarray"] == 0  # Hangul must lay the text out itself
    # every word of the Word file is there
    want = re.sub(r"\s+", "", docx_text(FIX / f"{stem}.docx"))
    got = re.sub(r"\s+", "", f["text"])
    assert not (Counter(want) - Counter(got)), "characters lost"


def test_s1_character_and_paragraph_formats(converted):
    f = features(converted["s1_text"][0].output)
    c = f["chars"]
    for k in ("bold", "italic", "underline", "strike", "sup", "sub", "color", "shade"):
        assert c[k] >= 1, k
    assert f["page_breaks"] == 1
    assert f["outline"] == 4  # heading 1 / heading 2 paragraphs keep their outline level
    head = etree.fromstring(zipfile.ZipFile(converted["s1_text"][0].output).read("Contents/header.xml"))
    faces = {e.get("face") for e in head.iter() if L(e.tag) == "font"}
    assert {"맑은 고딕", "바탕", "돋움"} <= faces


def test_s2_lists(converted):
    f = features(converted["s2_list"][0].output)
    assert f["bullets"] == 5 and f["numbers"] == 5


def test_s3_tables(converted):
    f = features(converted["s3_table"][0].output)
    assert f["tables"] == 4 and f["nested_table"] == 1 and f["merged"] == 2 and f["cell_fills"] == 4


def test_s4_page_features(converted):
    f = features(converted["s4_page"][0].output)
    assert f["sections"] == 3 and f["landscape"] == 1 and f["two_columns"] == 1
    assert f["header"] == 1 and f["footer"] == 1 and f["page_num"] == 1
    assert f["footnotes"] == 2 and f["hyperlinks"] == 1 and f["pictures"] == 1
    assert "첫째 각주의 설명 글입니다" in f["text"]


def test_s5_mixed(converted):
    f = features(converted["s5_mixed"][0].output)
    assert f["tables"] == 1 and f["pictures"] == 1 and f["footer"] == 1 and f["page_num"] == 1
    assert f["bullets"] == 4 and f["numbers"] == 3


# ---------------------------------------------------------------------------- refusals and bad input
def test_refuses_existing_output(tmp_path):
    dst = tmp_path / "a.hwpx"
    dst.write_bytes(b"keep me")
    r = convert(FIX / "s1_text.docx", dst)
    assert not r.ok and r.error and dst.read_bytes() == b"keep me"
    r = convert(FIX / "s1_text.docx", dst, overwrite=True)
    assert r.ok and dst.read_bytes()[:2] == b"PK"


def test_never_overwrites_source(tmp_path):
    src = tmp_path / "same.docx"
    src.write_bytes((FIX / "s1_text.docx").read_bytes())
    before = src.read_bytes()
    r = convert(src, src, overwrite=True)
    assert not r.ok and src.read_bytes() == before


def _korean(msg: str | None) -> bool:
    return bool(msg) and "\n" not in msg and re.search(r"[가-힣]", msg) is not None


@pytest.mark.parametrize("payload", [
    b"plain text, not a word file",
    b"PK\x03\x04broken zip",
    b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 512 + "EncryptionInfo".encode("utf-16-le"),
    b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 512,
])
def test_bad_input_is_a_korean_error(tmp_path, payload):
    src = tmp_path / "bad.docx"
    src.write_bytes(payload)
    r = convert(src, tmp_path / "bad.hwpx")
    assert not r.ok and _korean(r.error)
    assert not (tmp_path / "bad.hwpx").exists()


def test_zip_without_word_document(tmp_path):
    src = tmp_path / "x.docx"
    with zipfile.ZipFile(src, "w") as z:
        z.writestr("hello.txt", "hi")
    r = convert(src, tmp_path / "x.hwpx")
    assert not r.ok and _korean(r.error)


def test_missing_source(tmp_path):
    r = convert(tmp_path / "nope.docx", tmp_path / "nope.hwpx")
    assert not r.ok and _korean(r.error)


# ---------------------------------------------------------------------------- unsupported features
CT = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml"
 ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""
RELS = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
 Target="word/document.xml"/></Relationships>"""
DOC = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>{body}<w:sectPr>
<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"
 w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>"""


def make_docx(path: Path, body: str) -> Path:
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", DOC.format(body=body))
    return path


def test_unsupported_features_are_reported_and_text_kept(tmp_path):
    body = (
        '<w:p><w:r><w:t xml:space="preserve">앞 글 </w:t></w:r>'
        '<w:ins w:id="1" w:author="가"><w:r><w:t>넣은 글</w:t></w:r></w:ins>'
        '<w:del w:id="2" w:author="가"><w:r><w:delText>지운 글</w:delText></w:r></w:del>'
        '<w:commentRangeStart w:id="0"/><w:r><w:t>메모 단 글</w:t></w:r><w:commentRangeEnd w:id="0"/>'
        '<w:r><w:commentReference w:id="0"/></w:r></w:p>'
        '<w:p><m:oMath><m:r><m:t>x+1</m:t></m:r></m:oMath></w:p>'
        '<w:p><w:r><w:drawing><wp:anchor><wp:extent cx="914400" cy="457200"/><a:graphic>'
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
        '<wps:txbx><w:txbxContent><w:p><w:r><w:t>글상자 안 글</w:t></w:r></w:p></w:txbxContent></wps:txbx>'
        '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>'
        '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><a:graphic>'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/></a:graphic>'
        '</wp:inline></w:drawing></w:r></w:p>'
    )
    src = make_docx(tmp_path / "u.docx", body)
    r = convert(src, tmp_path / "u.hwpx")
    assert r.ok, r.error
    kinds = {l.feature: l for l in r.losses}
    for feature in ("tracked_change", "comment", "equation", "textbox", "chart"):
        assert feature in kinds, feature
        assert _korean(kinds[feature].label)
    text = features(r.output)["text"]
    for kept in ("앞 글", "넣은 글", "메모 단 글", "x+1", "글상자 안 글"):
        assert kept in text
    assert "지운 글" not in text
    assert hwpx.validate_editor_open_safety(str(r.output)).ok


def test_continuous_section_columns_stay_in_one_section(tmp_path):
    body = ('<w:p><w:r><w:t>한 단</w:t></w:r></w:p>'
            '<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440" w:header="720" w:footer="720"'
            ' w:gutter="0"/></w:sectPr></w:pPr></w:p>'
            '<w:p><w:r><w:t>두 단 글</w:t></w:r></w:p>'
            '<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:cols w:num="2" w:space="425"/>'
            '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440"'
            ' w:right="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:pPr></w:p>')
    src = make_docx(tmp_path / "c.docx", body)
    r = convert(src, tmp_path / "c.hwpx")
    assert r.ok, r.error
    f = features(r.output)
    assert f["sections"] == 1 and f["two_columns"] == 1


def test_large_document_is_fast(tmp_path):
    para = ('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>굵은 머리 </w:t></w:r><w:r><w:t xml:space="preserve">'
            + "가상의 긴 문장으로 쪽을 채웁니다. " * 12 + '</w:t></w:r></w:p>')
    row = '<w:tr>' + '<w:tc><w:p><w:r><w:t>칸</w:t></w:r></w:p></w:tc>' * 4 + '</w:tr>'
    table = '<w:tbl><w:tblGrid>' + '<w:gridCol w:w="2000"/>' * 4 + '</w:tblGrid>' + row * 20 + '</w:tbl>'
    body = (para * 40 + table) * 25  # about 100 pages
    src = make_docx(tmp_path / "big.docx", body)
    r = convert(src, tmp_path / "big.hwpx")
    assert r.ok, r.error
    assert r.seconds < 15
