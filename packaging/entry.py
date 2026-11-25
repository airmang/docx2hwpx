"""PyInstaller 진입점 (상대 import 가 있는 패키지를 직접 시작할 수 없어서 따로 둔다)."""
import sys

from docx2hwpx.__main__ import main

if __name__ == "__main__":
    sys.exit(main())
