# PyInstaller onedir: runtime Qt e DLLs do solver incluídos.
from PyInstaller.utils.hooks import collect_dynamic_libs, collect_data_files
from pathlib import Path
import os, sys
binaries = collect_dynamic_libs('ortools')
datas = [('assets', 'assets')] + collect_data_files('ortools')
a = Analysis(['main.py'], pathex=[SPECPATH], binaries=binaries, datas=datas,
             hiddenimports=['ortools.sat.python.cp_model', 'openpyxl'],
             excludes=['PySide6.QtWebEngineCore','PySide6.QtWebEngineWidgets',
                       'PySide6.QtPdf','IPython','ipykernel','matplotlib','pyscipopt'])
# Usa o ICU nativo do Windows e evita DLLs de ferramentas externas no PATH.
raizes = [Path(sys.prefix).resolve(), Path(sys.base_prefix).resolve(),
          Path(os.environ['SystemRoot']).resolve()]
a.binaries = [entry for entry in a.binaries
              if any(Path(entry[1]).resolve().is_relative_to(base) for base in raizes)
              and entry[0].lower() not in ('icuuc.dll', 'icudt78.dll')]
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name='UniSoma',
          debug=False, bootloader_ignore_signals=False, strip=False, upx=False, console=False)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name='UniSoma-2026')

