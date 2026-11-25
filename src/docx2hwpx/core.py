"""Converter core: one .docx in, one .hwpx out (see api.convert for the contract)."""
from __future__ import annotations

import os
import tempfile
import time
import zipfile
from pathlib import Path

from .api import ConversionResult, ProgressCallback
from .losses import LossLog

OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _check_input(source: Path) -> str | None:
    """A one-line Korean reason the file cannot be converted, or None."""
    if not source.exists():
        return "원본 파일을 찾을 수 없습니다."
    if not source.is_file():
        return "원본이 파일이 아닙니다."
    try:
        with open(source, "rb") as fh:
            head = fh.read(8)
            if head == OLE_MAGIC:
                fh.seek(0)
                blob = fh.read(1 << 20)
                if "EncryptionInfo".encode("utf-16-le") in blob or b"EncryptedPackage" in blob \
                        or "EncryptedPackage".encode("utf-16-le") in blob:
                    return "암호가 걸린 문서는 바꿀 수 없습니다. Word에서 암호를 푼 뒤 다시 시도해 주세요."
                return "예전 Word 형식(.doc)은 바꿀 수 없습니다. Word에서 .docx로 저장한 뒤 다시 시도해 주세요."
    except OSError:
        return "원본 파일을 읽을 수 없습니다."
    if not zipfile.is_zipfile(source):
        return "Word 문서(.docx)가 아니거나 파일이 손상되었습니다."
    return None


def _same_file(a: Path, b: Path) -> bool:
    try:
        if a.exists() and b.exists():
            return os.path.samefile(a, b)
    except OSError:
        pass
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def convert_file(source: Path, output: Path, *, overwrite: bool = False,
                 progress: ProgressCallback | None = None) -> ConversionResult:
    t0 = time.time()
    source, output = Path(source), Path(output)

    def fail(msg: str) -> ConversionResult:
        return ConversionResult(source=source, output=output, ok=False, error=msg, seconds=time.time() - t0)

    def report(frac: float, text: str) -> None:
        if progress:
            try:
                progress(frac, text)
            except Exception:  # noqa: BLE001 - a GUI callback must not break the conversion
                pass

    report(0.0, "준비 중")
    if _same_file(source, output):
        return fail("원본 파일에 덮어쓸 수 없습니다. 다른 이름으로 저장해 주세요.")
    if output.exists() and not overwrite:
        return fail("같은 이름의 파일이 이미 있습니다.")
    if output.exists() and output.is_dir():
        return fail("저장할 곳이 폴더입니다. 파일 이름을 정해 주세요.")
    reason = _check_input(source)
    if reason:
        return fail(reason)

    from .docx_reader import DocxError, DocxReader
    from .writer import HwpxWriter

    losses = LossLog()
    try:
        report(0.05, "Word 문서를 읽는 중")
        reader = DocxReader(str(source), losses)
        model = reader.read()
        report(0.3, "한글 문서로 옮기는 중")
        doc = HwpxWriter(model, losses, report).write()
        report(0.88, "저장 중")
        data = doc.to_bytes()
    except DocxError as ex:
        return fail(str(ex))
    except (zipfile.BadZipFile, KeyError, EOFError):
        return fail("Word 문서(.docx)가 손상되어 읽을 수 없습니다.")
    except MemoryError:
        return fail("문서가 너무 커서 바꿀 수 없습니다.")
    except Exception as ex:  # noqa: BLE001 - never raise to the GUI
        return fail(f"변환 중 오류가 났습니다({type(ex).__name__}).")

    try:
        output.parent.mkdir(parents=True, exist_ok=True)
        fd, tmp = tempfile.mkstemp(prefix=".~" + output.stem, suffix=".hwpx.tmp", dir=str(output.parent))
        try:
            with os.fdopen(fd, "wb") as fh:
                fh.write(data)
            if output.exists() and not overwrite:
                return fail("같은 이름의 파일이 이미 있습니다.")
            os.replace(tmp, output)
        finally:
            if os.path.exists(tmp):
                os.unlink(tmp)
    except PermissionError:
        return fail("파일을 저장할 수 없습니다. 같은 파일이 다른 프로그램(한글 등)에서 열려 있는지 확인해 주세요.")
    except OSError:
        return fail("파일을 저장할 수 없습니다. 저장할 곳을 확인해 주세요.")
    report(1.0, "끝")
    return ConversionResult(source=source, output=output, ok=True, losses=losses.to_list(),
                            seconds=time.time() - t0)
