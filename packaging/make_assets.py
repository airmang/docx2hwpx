"""app.ico 와 version_info.txt 를 만든다 (build.ps1 이 부른다)."""
from __future__ import annotations

import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "src"))


def _font(names, size):
    for n in names:
        try:
            return ImageFont.truetype(n, size)
        except OSError:
            continue
    return ImageFont.load_default()


def make_icon(path: Path) -> None:
    S = 512
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    # 왼쪽 워드 파랑, 오른쪽 한글 파랑 계열 둥근 배지
    d.rounded_rectangle((16, 16, S - 16, S - 16), radius=96, fill=(24, 90, 189, 255))
    d.rounded_rectangle((S // 2 + 8, 16, S - 16, S - 16), radius=96, fill=(0, 140, 140, 255))
    d.rectangle((S // 2 + 8, 16, S // 2 + 100, S - 16), fill=(0, 140, 140, 255))
    bold = ["malgunbd.ttf", "C:/Windows/Fonts/malgunbd.ttf", "arialbd.ttf"]
    w = _font(bold, 190)
    h = _font(bold, 200)
    d.text((S * 0.25, S * 0.5), "W", font=w, fill="white", anchor="mm")
    d.text((S * 0.75, S * 0.52), "한", font=h, fill="white", anchor="mm")
    # 가운데 화살표
    y = S // 2
    d.polygon([(S // 2 - 34, y - 26), (S // 2 + 6, y - 26), (S // 2 + 6, y - 52), (S // 2 + 58, y),
               (S // 2 + 6, y + 52), (S // 2 + 6, y + 26), (S // 2 - 34, y + 26)],
              fill=(255, 255, 255, 235))
    img.save(path, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])


def make_version(path: Path) -> None:
    from docx2hwpx import __version__

    parts = (__version__.split(".") + ["0", "0", "0"])[:3]
    nums = ", ".join(p for p in parts) + ", 0"
    text = f'''VSVersionInfo(
  ffi=FixedFileInfo(filevers=({nums}), prodvers=({nums}), mask=0x3f, flags=0x0, OS=0x40004, fileType=0x1, subtype=0x0, date=(0, 0)),
  kids=[
    StringFileInfo([StringTable('041204b0', [
      StringStruct('CompanyName', ''),
      StringStruct('FileDescription', '워드 → 한글 변환기'),
      StringStruct('FileVersion', '{__version__}'),
      StringStruct('InternalName', 'docx2hwpx'),
      StringStruct('OriginalFilename', '워드한글변환기.exe'),
      StringStruct('ProductName', '워드 → 한글 변환기'),
      StringStruct('ProductVersion', '{__version__}')])]),
    VarFileInfo([VarStruct('Translation', [0x0412, 1200])])
  ]
)
'''
    path.write_text(text, encoding="utf-8")


if __name__ == "__main__":
    make_icon(HERE / "app.ico")
    make_version(HERE / "version_info.txt")
    print("wrote app.ico, version_info.txt")
