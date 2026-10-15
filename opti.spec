# PyInstaller spec for the StudClassOpti desktop app.
#
# Build (same command on macOS and Windows CI):
#     pyinstaller opti.spec
#
# Produces a one-folder bundle under dist/ (onedir): reliable startup and easy
# to inspect which native libs got pulled in. OR-Tools ships native shared
# libraries + .proto data that PyInstaller's default analysis misses, so we
# collect it explicitly. PySide6 has built-in hooks and needs no help.

from PyInstaller.utils.hooks import collect_all

datas, binaries, hiddenimports = [], [], []
for pkg in ("ortools",):
    pkg_datas, pkg_binaries, pkg_hiddenimports = collect_all(pkg)
    datas += pkg_datas
    binaries += pkg_binaries
    hiddenimports += pkg_hiddenimports

a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
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
    name="opti",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,  # GUI app: no console window on Windows
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="opti",
)
