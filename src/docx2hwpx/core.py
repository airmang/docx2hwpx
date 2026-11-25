"""Converter core: placeholder until the core implementation lands. Keeps the GUI runnable meanwhile."""
from __future__ import annotations

import time
from pathlib import Path

from .api import ConversionResult, ProgressCallback


def convert_file(source: Path, output: Path, *, overwrite: bool = False,
                 progress: ProgressCallback | None = None) -> ConversionResult:
    t0 = time.time()
    if progress:
        progress(0.0, "준비 중")
    return ConversionResult(source=source, output=output, ok=False, error="변환 엔진이 아직 없습니다(개발 중).",
                            seconds=time.time() - t0)
