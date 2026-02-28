# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['process_acqua_reports.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\jianzou\\AppData\\Local\\Programs\\Python\\Python313\\Lib\\site-packages\\docx/templates', 'docx/templates')],
    hiddenimports=['docx', 'docx.document', 'docx.opc.constants', 'docx.opc.package', 'docx.opc.packuri', 'docx.opc.part', 'docx.opc.phys_pkg', 'docx.opc.rel', 'lxml._elementpath', 'lxml.etree', 'win32com', 'win32com.client', 'pythoncom', 'pywintypes'],
    hookspath=[],
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
    a.binaries,
    a.datas,
    [],
    name='ACQUA_ReportReviewer_v1.2.0',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
