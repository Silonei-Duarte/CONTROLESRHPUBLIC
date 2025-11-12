import os
from glob import glob
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# inclui todos arquivos de Essenciais/
essenciais_files = [(f, '.') for f in glob('Essenciais/*') if os.path.isfile(f)]

# caminho completo do kaleido
kaleido_path = r'C:\Users\silonei.duarte.BRUNO\AppData\Local\Programs\Python\Python313\Lib\site-packages\kaleido'

# adiciona conteúdo do kaleido e dependências
datas = (
    essenciais_files
    + collect_data_files('plotly', include_py_files=False)
    + collect_data_files('kaleido', include_py_files=False)
    + [(kaleido_path, 'kaleido')]
    + [(os.path.join(kaleido_path, 'executable', 'bin', 'kaleido.exe'), 'kaleido/executable/bin')]
    + [(os.path.join(os.path.dirname(kaleido_path), 'kaleido-0.2.1.dist-info'), 'kaleido-0.2.1.dist-info')]
)


hiddenimports = (
    collect_submodules('plotly')
    + collect_submodules('kaleido')
    + collect_submodules('fontTools')
    + collect_submodules('cryptography')
    + collect_submodules('oracledb')
)

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=['hooks'],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='main',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['icone.ico'],
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='main',
)
