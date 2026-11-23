# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
ROOT = Path(SPECPATH)
datas = []
for folder in ("templates", "data"):
    source = ROOT / folder
    if source.exists():
        datas.append((str(source), folder))
a = Analysis(["main.py"], pathex=[str(ROOT)], binaries=[], datas=datas,
             hiddenimports=[], hookspath=[], hooksconfig={}, runtime_hooks=[],
             excludes=["pytest", "unittest"], noarchive=False)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name="AptechMonthlyBSR",
          debug=False, bootloader_ignore_signals=False, strip=False, upx=True,
          console=False, disable_windowed_traceback=False)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=True,
               name="AptechMonthlyBSR")
