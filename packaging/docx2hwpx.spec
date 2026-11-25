# -*- mode: python ; coding: utf-8 -*-
# 사용: pyinstaller packaging/docx2hwpx.spec  (build.ps1 이 대신 실행한다)
import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules, collect_dynamic_libs

HERE = os.path.dirname(os.path.abspath(SPEC))
ROOT = os.path.dirname(HERE)

datas, binaries, hiddenimports = [], [], []
for pkg in ("hwpx", "docx", "lxml", "tkinterdnd2"):
    try:
        datas += collect_data_files(pkg)
        hiddenimports += collect_submodules(pkg)
    except Exception as e:  # 없는 패키지는 건너뜀
        print("collect skip", pkg, e)
binaries += collect_dynamic_libs("tkinterdnd2")
hiddenimports += collect_submodules("docx2hwpx") + collect_submodules("PIL")
datas += [(os.path.join(HERE, "app.ico"), ".")]

a = Analysis(
    [os.path.join(HERE, "entry.py")],
    pathex=[os.path.join(ROOT, "src")],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    excludes=["pytest", "numpy", "matplotlib", "IPython", "unittest", "pydoc_data"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name="워드한글변환기",
    console=False,
    icon=os.path.join(HERE, "app.ico"),
    version=os.path.join(HERE, "version_info.txt"),
    upx=False,
    runtime_tmpdir=None,
)
