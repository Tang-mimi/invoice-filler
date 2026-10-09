# -*- mode: python ; coding: utf-8 -*-
import os
import sys

from PyInstaller.utils.hooks import collect_all

# ---- 版本号的唯一来源：invoice_filler/__init__.py ----
# 打包产物名与 exe 文件属性都从这里取，避免"改了代码忘了改打包脚本"。
_BASE = globals().get("SPECPATH") or os.getcwd()
if _BASE not in sys.path:
    sys.path.insert(0, _BASE)
from invoice_filler import __version__  # noqa: E402

APP_NAME = "发票填表工具"
OUT_NAME = f"{APP_NAME}-v{__version__}"


def _ver_tuple(v):
    parts = []
    for x in str(v).split("."):
        try:
            parts.append(int(x))
        except ValueError:
            parts.append(0)
    while len(parts) < 4:
        parts.append(0)
    return tuple(parts[:4])


# ---- 写进 exe 文件属性（右键→属性→详细信息 里的「文件版本 / 产品版本」）----
try:
    from PyInstaller.utils.win32.versioninfo import (
        FixedFileInfo, StringFileInfo, StringStruct, StringTable,
        VarFileInfo, VarStruct, VSVersionInfo)

    _V = _ver_tuple(__version__)
    version_info = VSVersionInfo(
        ffi=FixedFileInfo(filevers=_V, prodvers=_V, mask=0x3F, flags=0x0,
                          OS=0x40004, fileType=0x1, subtype=0x0, date=(0, 0)),
        kids=[
            StringFileInfo([StringTable("080404B0", [
                StringStruct("CompanyName", ""),
                StringStruct("FileDescription", "发票填表工具 - 数电票批量识别填表"),
                StringStruct("FileVersion", __version__),
                StringStruct("InternalName", APP_NAME),
                StringStruct("OriginalFilename", f"{OUT_NAME}.exe"),
                StringStruct("ProductName", APP_NAME),
                StringStruct("ProductVersion", __version__),
            ])]),
            VarFileInfo([VarStruct("Translation", [2052, 1200])]),
        ],
    )
except Exception:      # 拿不到也不能让打包失败，只是少了文件属性
    version_info = None

datas = [('invoice_filler/builtin_template1.json', 'invoice_filler'),
         ('invoice_filler/builtin_template.xlsx', 'invoice_filler')]
binaries = []
hiddenimports = []
tmp_ret = collect_all('rapidocr_onnxruntime')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('tkinterdnd2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    name=OUT_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version=version_info,
)
