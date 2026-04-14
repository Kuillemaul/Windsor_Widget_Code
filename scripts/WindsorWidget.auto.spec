# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
from PyInstaller.utils.hooks import collect_submodules

builder_root = Path(SPECPATH).parent
stage_dir = builder_root / "staged_app"
dist_dir = builder_root / "dist"
work_dir = builder_root / "build-temp"

hiddenimports = []
hiddenimports += collect_submodules("PySide6.QtCharts")

datas = []
review_module = stage_dir / "yu_order_review_export_test_window.py"
if review_module.exists():
    datas.append((str(review_module), "."))

assets_dir = stage_dir / "assets"
if assets_dir.exists():
    datas.append((str(assets_dir), "assets"))

icon_path = builder_root / "assets.ico"
icon_value = str(icon_path) if icon_path.exists() else None

a = Analysis(
    [str(stage_dir / "main_patched_status_yu.py")],
    pathex=[str(stage_dir)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="WindsorWidget",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon=icon_value,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="WindsorWidget",
)
