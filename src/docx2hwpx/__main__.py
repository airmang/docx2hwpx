"""`python -m docx2hwpx` — 인자가 없으면 창을 열고, 파일을 주면 명령줄에서 바로 변환한다."""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

from . import __version__
from .api import convert


def _fix_streams() -> None:
    """콘솔 없는(windowed) exe 에서는 stdout 이 None 이다. 한글이 깨지는 콘솔에서도 죽지 않게 한다."""
    for name in ("stdout", "stderr"):
        stream = getattr(sys, name)
        if stream is None:
            setattr(sys, name, open(os.devnull, "w", encoding="utf-8"))
        else:
            try:
                stream.reconfigure(errors="replace")
            except Exception:
                pass


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="docx2hwpx",
        description="워드(.docx) 파일을 한글(.hwpx) 파일로 바꿉니다. 인자 없이 실행하면 창이 열립니다.")
    p.add_argument("inputs", nargs="*", metavar="입력.docx", help="변환할 워드 파일")
    p.add_argument("-o", "--output", metavar="출력.hwpx", help="저장할 파일 이름 (입력이 하나일 때만)")
    p.add_argument("--overwrite", action="store_true", help="같은 이름의 .hwpx 가 있으면 덮어씀")
    p.add_argument("--version", action="store_true", help="버전을 보여 주고 끝냄")
    p.add_argument("--selftest", metavar="결과파일", help=argparse.SUPPRESS)
    p.add_argument("--selftest-docx", metavar="입력.docx", help=argparse.SUPPRESS)
    return p


def _print_result(res) -> None:
    if res.ok:
        print(f"완료: {res.source} -> {res.output}  ({res.seconds:.1f}초)")
        if res.losses:
            print("  한글로 옮기지 못했거나 달라진 부분:")
            for loss in res.losses:
                mark = "빠짐" if loss.kind == "dropped" else "달라짐"
                print(f"   - [{mark}] {loss.label}")
        else:
            print("  빠지거나 달라진 부분 없음")
    else:
        print(f"실패: {res.source}\n  {res.error}")


def run_cli(args: argparse.Namespace) -> int:
    from .gui import expand_inputs, unique_output_path

    inputs = [Path(x) for x in args.inputs]
    missing = [p for p in inputs if not p.exists()]
    for p in missing:
        print(f"찾을 수 없습니다: {p}")
    if missing:
        return 2
    if args.output and len(inputs) != 1:
        print("-o 는 입력 파일이 하나일 때만 쓸 수 있습니다.")
        return 2
    if args.output:
        sources = [inputs[0]]
    else:
        sources = expand_inputs(inputs)
        if not sources:
            print("변환할 .docx 파일이 없습니다.")
            return 2
    bad = 0
    reserved: list[Path] = []
    for src in sources:
        if args.output:
            out = Path(args.output)
        elif args.overwrite:
            out = src.with_suffix(".hwpx")
        else:
            out = unique_output_path(src, None, reserved)
            reserved.append(out)

        def progress(frac: float, text: str, name=src.name) -> None:
            if text:
                print(f"  [{int(frac * 100):3d}%] {text}")

        res = convert(src, out, overwrite=args.overwrite, progress=progress if len(sources) == 1 else None)
        _print_result(res)
        bad += 0 if res.ok else 1
    return 1 if bad else 0


def selftest(result_path: str, docx: str | None = None) -> int:
    """창을 만들었다 닫고, 변환 엔진과 의존 모듈이 불러와지는지 확인해 결과를 파일에 적는다."""
    lines: list[str] = []
    ok = True
    try:
        import lxml.etree  # noqa: F401
        import PIL.Image  # noqa: F401
        import docx2hwpx.core  # noqa: F401
        import hwpx  # noqa: F401
        lines.append("imports: OK")
        import importlib
        import pkgutil

        bad = []
        count = 0
        for pkg in (hwpx, ):
            for m in pkgutil.walk_packages(pkg.__path__, pkg.__name__ + "."):
                try:
                    importlib.import_module(m.name)
                    count += 1
                except Exception as e:  # noqa: BLE001
                    bad.append(f"{m.name}: {e!r}")
        lines.append(f"hwpx submodules: {count} imported, {len(bad)} failed")
        lines.extend("  " + b for b in bad)
        from . import gui

        root = gui._make_root()
        root.withdraw()
        app = gui.App(root)
        root.update()
        dnd = "yes" if app.dnd_ok else "no"
        root.destroy()
        lines.append(f"window: OK (drag-drop {dnd})")
        if docx:
            import tempfile

            with tempfile.TemporaryDirectory() as td:
                res = convert(docx, Path(td) / "t.hwpx")
                lines.append(f"convert: {'OK' if res.ok else 'FAIL'} {res.error or ''} losses={len(res.losses)}")
                ok = ok and res.ok
    except Exception as e:  # noqa: BLE001
        import traceback

        ok = False
        lines.append("ERROR: " + repr(e))
        lines.append(traceback.format_exc())
    lines.append("OK" if ok else "FAILED")
    Path(result_path).write_text("\n".join(lines) + "\n", encoding="utf-8")
    return 0 if ok else 1


def main(argv: list[str] | None = None) -> int:
    _fix_streams()
    args = _build_parser().parse_args(argv)
    if args.version:
        print(f"워드 → 한글 변환기 {__version__}")
        return 0
    if args.selftest:
        return selftest(args.selftest, args.selftest_docx)
    if args.inputs:
        return run_cli(args)
    from .gui import main as gui_main

    gui_main()
    return 0


if __name__ == "__main__":
    sys.exit(main())
