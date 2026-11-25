"""Public contract shared by the converter core and the GUI. Do not change field names without updating both."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Literal

LossKind = Literal["dropped", "changed"]


@dataclass(frozen=True)
class Loss:
    """Something the Hangul copy does not carry as Word had it.

    feature: short stable id ("textbox", "equation", "comment", "floating_picture", ...)
    kind: "dropped" (gone) or "changed" (there, but different)
    count: how many places
    label: Korean text for the person ("글상자 3개가 빠졌습니다")
    """

    feature: str
    kind: LossKind
    count: int
    label: str


@dataclass
class ConversionResult:
    source: Path
    output: Path
    ok: bool
    losses: list[Loss] = field(default_factory=list)
    error: str | None = None  # Korean, one line, when ok is False
    seconds: float = 0.0
    pages_hint: int | None = None  # None: unknown


class ConversionError(Exception):
    """A .docx that cannot be converted at all (not a Word file, encrypted, broken zip)."""


ProgressCallback = Callable[[float, str], None]  # (0.0-1.0, Korean status text)


def convert(source: Path | str, output: Path | str, *, overwrite: bool = False,
            progress: ProgressCallback | None = None) -> ConversionResult:
    """Convert one .docx to .hwpx. Never raises for a bad input file: returns ok=False with error.

    Refuses to overwrite an existing output unless overwrite=True (never overwrite the source).
    """
    from .core import convert_file  # the core lives in core.py; imported lazily so the GUI starts fast

    return convert_file(Path(source), Path(output), overwrite=overwrite, progress=progress)
