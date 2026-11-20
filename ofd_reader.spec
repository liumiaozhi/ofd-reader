# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置 —— OFD 阅读转换器

用法（在本目录下）：
    python -m PyInstaller ofd_reader.spec --noconfirm --clean

说明：
  · 目标运行环境是 Windows 7 SP1 及以上（含 32/64 位），
    因此用 Python 3.8 打包——它是最后一个支持 Win7 的官方版本。
  · 本 spec 同时兼容 PyInstaller 4.x 与 6.x：
      - 4.x 是 Windows 7 上验证最充分的版本（若 6.x 打出的包在 Win7 上
        报 “Failed to execute script”，可 pip install pyinstaller==4.10
        后重新执行本 spec，无需改动任何内容）。
      - 6.x 是当前 Python 3.8 环境里默认安装的版本。
  · 产物：dist/ofd阅读转换器.exe（单文件，免安装）
"""

import os

from PyInstaller import __version__ as PYI_VERSION

SPEC_DIR = SPECPATH          # noqa: F821  —— PyInstaller 注入
APP_NAME = "ofd阅读转换器"
PYI_MAJOR = int(str(PYI_VERSION).split(".")[0])

datas = [
    (os.path.join(SPEC_DIR, "assets", "logo.ico"), "assets"),
    (os.path.join(SPEC_DIR, "assets", "logo.png"), "assets"),
]

hiddenimports = [
    "docx",
    "lxml.etree",
    "lxml._elementpath",
    "PIL._tkinter_finder",
    "PIL.ImageTk",
    "PIL.ImageFont",
    "PIL.PdfImagePlugin",
    "PIL.IcoImagePlugin",
    "PIL.PngImagePlugin",
    "PIL.JpegImagePlugin",
    "PIL.BmpImagePlugin",
    "PIL.GifImagePlugin",
    "PIL.TiffImagePlugin",
]

# 这些库在本机里装着，但程序完全用不到；排掉能显著减小体积
excludes = [
    "numpy", "pandas", "scipy", "matplotlib", "IPython", "jupyter",
    "PySide2", "PySide6", "PyQt5", "PyQt6", "notebook", "pytest", "sphinx",
]

analysis_kwargs = dict(
    pathex=[SPEC_DIR],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
)
if PYI_MAJOR >= 5:
    analysis_kwargs["hooksconfig"] = {}
if PYI_MAJOR >= 6:
    analysis_kwargs["optimize"] = 0

a = Analysis(["ofd_reader.py"], **analysis_kwargs)      # noqa: F821

if PYI_MAJOR >= 6:
    pyz = PYZ(a.pure)                                    # noqa: F821
    extra = a.datas
else:                                                    # PyInstaller 4.x / 5.x
    pyz = PYZ(a.pure, a.zipped_data, cipher=None)        # noqa: F821
    extra = a.zipfiles + a.datas

exe_kwargs = dict(
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,           # 窗口程序，不弹黑框
    icon=os.path.join(SPEC_DIR, "assets", "logo.ico"),
    version=os.path.join(SPEC_DIR, "version_info.txt"),
)
if PYI_MAJOR >= 6:
    exe_kwargs["disable_windowed_traceback"] = False

# 单文件模式：把二进制、纯 Python 归档与数据文件全部塞进同一个 exe
exe = EXE(pyz, a.scripts, a.binaries, extra, [], **exe_kwargs)       # noqa: F821
