# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec — ConvertidorBancos v5.1 (edición completa CON Descarga Masiva SAT)
# ERSL Consultores 2026
#
# ⚠ CANAL PORTAL (playwright) — NO se congela en el .exe.
#   Playwright necesita binarios de navegador (~150-300 MB) que se instalan aparte
#   con `playwright install chromium` y que PyInstaller no empaqueta de forma
#   confiable. El canal Portal es un RESPALDO OPCIONAL; el canal PRINCIPAL es el
#   Web Service (satcfdi), que SÍ queda bundleado. En el .exe, si playwright/chromium
#   no están, la app degrada por su guardia de import (no rompe el resto).
#   Si en el futuro se quiere el Portal dentro del .exe, hay que quitar 'playwright'
#   de `excludes`, agregar collect_all('playwright') y una estrategia para los
#   browsers (p. ej. PLAYWRIGHT_BROWSERS_PATH junto al .exe) — y PROBARLO en real.
#
# ⚠ Este .spec necesita una corrida real de PyInstaller para confirmar el set de
#   hidden imports de satcfdi (arrastra lxml/cryptography/requests). Parsear != build.

import os
from PyInstaller.utils.hooks import collect_all, collect_submodules

block_cipher = None

# ── Plantilla con macro VBA (papel de trabajo) — se empaqueta ya firmada ──────
_PLANTILLA = 'plantilla_papel_trabajo.xlsm'
if not os.path.exists(_PLANTILLA):
    raise SystemExit(
        f"Falta {_PLANTILLA}. Genera la plantilla con macro una vez:\n"
        "   pip install pywin32 && python construir_plantilla.py\n"
        "y FIRMALA (Alt+F11 > Tools > Digital Signature > ERSL Consultores) antes\n"
        "de empaquetar. El build NO la regenera para no borrar la firma del VBA."
    )
datas_plantilla = [(_PLANTILLA, '.')]

datas_plumber,  binaries_plumber,  hiddens_plumber  = collect_all('pdfplumber')
datas_miner,    binaries_miner,    hiddens_miner    = collect_all('pdfminer')
datas_openpyxl, binaries_openpyxl, hiddens_openpyxl = collect_all('openpyxl')
datas_pil,      binaries_pil,      hiddens_pil      = collect_all('PIL')
# SAT Web Service (canal principal) + sus dependencias:
datas_sat,      binaries_sat,      hiddens_sat      = collect_all('satcfdi')
datas_lxml,     binaries_lxml,     hiddens_lxml     = collect_all('lxml')
datas_req,      binaries_req,      hiddens_req      = collect_all('requests')
datas_crypto,   binaries_crypto,   hiddens_crypto   = collect_all('cryptography')
# Auto-reparación de selectores del Portal (pure-python, chico):
datas_anth,     binaries_anth,     hiddens_anth     = collect_all('anthropic')
# UI rediseñado (temas .json) + BD Firebird de conciliacion + PDF de reportes:
datas_ctk,      binaries_ctk,      hiddens_ctk      = collect_all('customtkinter')
datas_fbdrv,    binaries_fbdrv,    hiddens_fbdrv    = collect_all('firebird.driver')
datas_fbbase,   binaries_fbbase,   hiddens_fbbase   = collect_all('firebird.base')
datas_rl,       binaries_rl,       hiddens_rl       = collect_all('reportlab')

a = Analysis(
    ['ConvertidorSantander.py'],
    pathex=['.'],
    binaries=(binaries_plumber + binaries_miner + binaries_openpyxl + binaries_pil +
              binaries_sat + binaries_lxml + binaries_req + binaries_crypto + binaries_anth +
              binaries_ctk + binaries_fbdrv + binaries_fbbase + binaries_rl),
    datas=(datas_plantilla + datas_plumber + datas_miner + datas_openpyxl + datas_pil +
           datas_sat + datas_lxml + datas_req + datas_crypto + datas_anth +
           datas_ctk + datas_fbdrv + datas_fbbase + datas_rl),
    hiddenimports=(
        hiddens_plumber + hiddens_miner + hiddens_openpyxl + hiddens_pil +
        hiddens_sat + hiddens_lxml + hiddens_req + hiddens_crypto + hiddens_anth +
        hiddens_ctk + hiddens_fbdrv + hiddens_fbbase + hiddens_rl +
        ['customtkinter', 'firebird.driver', 'firebird.base', 'reportlab'] +
        collect_submodules('pdfplumber') +
        collect_submodules('pdfminer') +
        collect_submodules('openpyxl') +
        collect_submodules('PIL') +
        collect_submodules('bancos') +
        collect_submodules('satcfdi') +
        collect_submodules('sat_saldos') +
        [
            'tkinter', 'tkinter.ttk', 'tkinter.filedialog', 'tkinter.messagebox',
            'tkinter.font', 'tkinter.scrolledtext', 'tkinter.simpledialog',
            'zlib', 'struct', 'hashlib', 'hmac', 'base64', 'uuid', 'json', 'pathlib',
            'socket', 'subprocess', 'sqlite3', 'smtplib', 'email', 'ssl',
            # módulo SAT Web Service local (import guardado con try/except):
            'sat_descarga',
        ]
    ),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'numpy', 'scipy', 'pandas', 'pytest',
        'jupyter', 'notebook', 'IPython',
        # playwright NO se congela (browsers aparte); canal Portal = respaldo opcional.
        # 'sat_descarga_portal' importa playwright; su import está guardado con
        # try/except en la app, así que su ausencia degrada el canal Portal sin romper.
        'playwright',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ConvertidorBancos',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,          # UPX dispara falsos positivos de AV/Defender. Sin comprimir.
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.ico',
    version='version_info.txt',   # metadata de editor/producto en el .exe
    manifest='app.manifest',      # DPI Per-Monitor-V2 + asInvoker
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,          # sin UPX tambien en el COLLECT (mismos falsos positivos)
    upx_exclude=[],
    name='ConvertidorBancos',
)
