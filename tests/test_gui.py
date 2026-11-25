"""GUI 도우미, 명령줄 처리, 창 만들기 시험."""
from __future__ import annotations

from pathlib import Path

import pytest

from docx2hwpx import __main__ as cli
from docx2hwpx.api import ConversionResult, Loss
from docx2hwpx.gui import (ST_FAIL, ST_OK, ST_PART, detail_text, expand_inputs, is_word_file,
                           status_for, unique_output_path)

FIX = Path(__file__).parent / "fixtures"


def test_unique_output_path_numbering(tmp_path):
    src = tmp_path / "보고서.docx"
    src.write_bytes(b"x")
    assert unique_output_path(src) == tmp_path / "보고서.hwpx"
    (tmp_path / "보고서.hwpx").write_bytes(b"x")
    assert unique_output_path(src) == tmp_path / "보고서 (1).hwpx"
    (tmp_path / "보고서 (1).hwpx").write_bytes(b"x")
    assert unique_output_path(src) == tmp_path / "보고서 (2).hwpx"


def test_unique_output_path_other_folder_and_reserved(tmp_path):
    a = tmp_path / "a" / "x.docx"
    b = tmp_path / "b" / "x.docx"
    out = tmp_path / "out"
    out.mkdir()
    first = unique_output_path(a, out)
    assert first == out / "x.hwpx"
    second = unique_output_path(b, out, reserved=[first])
    assert second == out / "x (1).hwpx"


def test_expand_inputs_skips_lock_files_and_others(tmp_path):
    (tmp_path / "a.docx").write_bytes(b"x")
    (tmp_path / "~$a.docx").write_bytes(b"x")
    (tmp_path / "b.DOCX").write_bytes(b"x")
    (tmp_path / "c.txt").write_bytes(b"x")
    (tmp_path / "old.doc").write_bytes(b"x")
    sub = tmp_path / "sub"
    sub.mkdir()
    (sub / "d.docx").write_bytes(b"x")
    (sub / "~$d.docx").write_bytes(b"x")
    got = expand_inputs([tmp_path, tmp_path / "a.docx", tmp_path / "missing.docx"])
    names = [p.name for p in got]
    assert sorted(names) == ["a.docx", "b.DOCX", "d.docx"]  # 중복 없음, 잠금 파일 없음
    assert not is_word_file(tmp_path / "~$a.docx")


def test_status_and_detail():
    src, out = Path("a.docx"), Path("a.hwpx")
    ok = ConversionResult(src, out, ok=True)
    part = ConversionResult(src, out, ok=True, losses=[Loss("textbox", "dropped", 2, "글상자 2개가 빠졌습니다")])
    bad = ConversionResult(src, out, ok=False, error="암호가 걸린 파일입니다.")
    assert status_for(ok) == ST_OK and status_for(part) == ST_PART and status_for(bad) == ST_FAIL
    assert "글상자 2개가 빠졌습니다" in detail_text(part, ST_PART)
    assert "암호가 걸린 파일입니다." in detail_text(bad, ST_FAIL)


def test_cli_version(capsys):
    assert cli.main(["--version"]) == 0
    assert "0.1.0" in capsys.readouterr().out


def test_cli_converts_with_unique_name_and_reports_losses(tmp_path, monkeypatch, capsys):
    src = tmp_path / "in.docx"
    src.write_bytes(b"x")
    (tmp_path / "in.hwpx").write_bytes(b"old")
    calls = []

    def fake(source, output, *, overwrite=False, progress=None):
        calls.append((Path(source), Path(output), overwrite))
        return ConversionResult(Path(source), Path(output), ok=True,
                                losses=[Loss("equation", "changed", 1, "수식 1개가 그림으로 바뀌었습니다")])

    monkeypatch.setattr(cli, "convert", fake)
    assert cli.main([str(src)]) == 0
    assert calls[0][1] == tmp_path / "in (1).hwpx" and calls[0][2] is False
    assert "수식 1개가 그림으로 바뀌었습니다" in capsys.readouterr().out


def test_cli_output_and_overwrite_flags(tmp_path, monkeypatch):
    src = tmp_path / "in.docx"
    src.write_bytes(b"x")
    seen = {}

    def fake(source, output, *, overwrite=False, progress=None):
        seen.update(out=Path(output), overwrite=overwrite)
        return ConversionResult(Path(source), Path(output), ok=False, error="안 됨")

    monkeypatch.setattr(cli, "convert", fake)
    assert cli.main([str(src), "-o", str(tmp_path / "x.hwpx"), "--overwrite"]) == 1  # 실패면 종료 코드 1
    assert seen == {"out": tmp_path / "x.hwpx", "overwrite": True}


def test_cli_errors(tmp_path, capsys):
    assert cli.main([str(tmp_path / "none.docx")]) == 2
    a = tmp_path / "a.docx"
    b = tmp_path / "b.docx"
    a.write_bytes(b"x")
    b.write_bytes(b"x")
    assert cli.main([str(a), str(b), "-o", "z.hwpx"]) == 2


def _display_root():
    import tkinter as tk
    try:
        from tkinterdnd2 import TkinterDnD
        return TkinterDnD.Tk()
    except Exception:
        try:
            return tk.Tk()
        except tk.TclError:
            pytest.skip("화면(display)이 없습니다")


def test_window_builds_adds_files_and_closes(tmp_path):
    from docx2hwpx.gui import App

    root = _display_root()
    root.withdraw()
    try:
        app = App(root)
        root.update()
        assert root.title() == "워드 → 한글 변환기"
        n = app.add_paths([str(FIX)])
        assert n == len(list(FIX.glob("*.docx")))
        assert app.add_paths([str(FIX)]) == 0  # 같은 파일은 다시 넣지 않음
        assert {i.status for i in app.items.values()} == {"대기"}
        first = app.order[0]
        app.tree.selection_set(first)
        app.remove_selected()
        assert len(app.items) == n - 1
        app.clear_all()
        assert not app.items
    finally:
        root.destroy()


def test_window_converts_in_worker_thread(tmp_path, monkeypatch):
    from docx2hwpx import gui

    def fake(source, output, *, overwrite=False, progress=None):
        if progress:
            progress(0.5, "중간")
        return ConversionResult(Path(source), Path(output), ok=True)

    monkeypatch.setattr(gui, "convert", fake)
    root = _display_root()
    root.withdraw()
    try:
        app = gui.App(root)
        app.add_paths([str(FIX / "s1_text.docx"), str(FIX / "s2_list.docx")])
        app.start()
        for _ in range(200):
            root.update()
            if not app.running:
                break
            root.after(20)
        assert [i.status for i in app.items.values()] == ["완료", "완료"]
    finally:
        root.destroy()
